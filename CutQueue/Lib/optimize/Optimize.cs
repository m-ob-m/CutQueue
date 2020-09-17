using CutQueue.Lib;
using CutQueue.Lib.Fabplan;
using CutQueue.Logging;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace CutQueue
{
    /// <summary>
    /// A class that optimizes batches from Fabplan using CutRite to create machining programs
    /// </summary>
    class Optimize
    {
        private static bool _inProgress = false;
        private enum ContinueStatus {No, Yes}

        /// <summary>
        /// Main constructor
        /// </summary>
        public Optimize()
        {}

        /// <summary>
        /// Optimizes batches until there is nothing left to optimize.
        /// </summary>
        public async Task DoOptimize()
        {
            if (!_inProgress)
            {
                _inProgress = true;
                ContinueStatus continueStatus = ContinueStatus.Yes;

                while (continueStatus == ContinueStatus.Yes)
                {
                    try
                    {
                        continueStatus = await OptimizeNextBatch();
                    }
                    catch (Exception e)
                    {
                        _inProgress = false;
                        throw new Exception("Could not optimize batch.", e);
                    }
                }

                _inProgress = false;
            }
        }

        /// <summary>
        /// Optimizes the next batch.
        /// </summary>
        /// <exception cref="Exception">Thrown when the server fails to return the information on the next job to optimize.</exception>
        /// <returns>A <c>ContinueStatus</c> value that tells the main optimization loop when if it should continue looping.</returns>
        private async Task<ContinueStatus> OptimizeNextBatch()
        {
            dynamic batch = null;
            try
            {
                batch = await GetNextBatchToOptimize();
            }
            catch (Exception e)
            {
                throw new Exception("Cannot retrieve next batch to optimize.", e);
            }

            if (batch != null)
            {
                try
                {
                    await OptimizeBatch(batch);
                }
                catch (Exception e)
                {
                    await UpdateEtatMpr(((dynamic)batch).id, 'E', e.Message);
                }
            }
            else
            {
                return ContinueStatus.No;
            }

            return ContinueStatus.Yes;
        }

        /// <summary>
        /// Queries Fabplan and returns the next batch to optimize.
        /// </summary>
        /// <exception cref="Exception">Thrown when the server fails to repond or responds improperly.</exception>
        /// <returns>An object that represents the next batch to optimize.</returns>
        private async Task<dynamic> GetNextBatchToOptimize()
        {
            UriBuilder builder = new UriBuilder()
            {
                Host = ConfigINI.Items["FABPLAN_HOST_NAME"].ToString(),
                Path = ConfigINI.Items["FABPLAN_GET_NEXT_BATCH_TO_OPTIMIZE_URL"].ToString(),
                Port = -1,
                Scheme = "http"
            };

            JObject rawBatch;
            try
            {
                rawBatch = await FabplanHttpRequest.Get(builder.ToString());
            }
            catch (Exception e)
            {
                throw new Exception("Failed retrieving next batch to optimize", e);
            }

            dynamic batch = new ExpandoObject();
            if (rawBatch != null)
            {
                try
                {
                    batch.id = rawBatch.GetValue("id").ToObject<long>();
                }
                catch (Exception e)
                {
                    throw new Exception("Next batch to optimize is missing its \"id\" member.", e);
                }

                try
                {
                    batch.name = rawBatch.GetValue("name").ToObject<string>();
                }
                catch (Exception e)
                {
                    throw new Exception("Next batch to optimize is missing its \"name\" member.", e);
                }

                try
                {
                    batch.pannels = rawBatch.GetValue("pannels").ToObject<string>();
                }
                catch (Exception e)
                {
                    throw new Exception("Next batch to optimize is missing its \"pannels\" member.", e);
                }
            }
            else
            {
                return null;
            }

            return batch;
        }

        /// <summary>
        /// Executes a process.
        /// </summary>
        /// <param name="procName">The filepath of the process to execute</param>
        /// <param name="arguments">A string of arguments used to start the process</param>
        /// <returns>The process' output</returns>
        private (int exitCode, string standardOutput) ExecuteProcess(string procName, string arguments)
        {
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo
            {
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden,
                FileName = procName,
                Arguments = arguments,
                RedirectStandardOutput = true,
                UseShellExecute = false
            };

            System.Diagnostics.Process process = new System.Diagnostics.Process
            {
                StartInfo = startInfo
            };
            process.Start();

            string standardOutput = process.StandardOutput.ReadToEnd();
            

            process.WaitForExit();

            return (process.ExitCode, standardOutput);
        }

        /// <summary>
        /// Optimizes a given batch using Cut Rite.
        /// </summary>
        /// <param name="batch"> The name of the batch to optimize</param>
        /// <exception cref="ArgumentException">Thrown when optimise.exe fails.</exception>
        private async Task OptimizeBatch(dynamic batch)
        {
            // Changer l'état pour en cours (P pour In Progress)
            await UpdateEtatMpr(((dynamic)batch).id, 'C');

            string batchName = ((dynamic)batch).name;      // Nom du fichier CSV

            int exitCode;
            string standardOutput;

            DeletePartListFile(batchName);

            // Import du CSV
            (exitCode, standardOutput) = ExecuteProcess(new Uri(new Uri($"{Application.StartupPath}/"), "autoit/importCSV.exe").LocalPath, $"\"{batchName}.txt\"");
            if (exitCode != 0)
            {
                throw new Exception(standardOutput);
            }

            // Génération du fichier des panneaux (.brd)
            (exitCode, standardOutput) = ExecuteProcess(new Uri(new Uri($"{Application.StartupPath}/"), "autoit/panneaux.exe").LocalPath, $"\"{batchName}.txt\"");
            if (exitCode != 0)
            {
                throw new Exception(standardOutput);
            }

            // Modification du fichier .brd pour avoir les bons panneaux
            ModifyPanneaux(((dynamic)batch).pannels, ((dynamic)batch).name);

            // Optimisation
            (exitCode, standardOutput) = ExecuteProcess(new Uri(new Uri($"{Application.StartupPath}/"), "autoit/optimise.exe").LocalPath, $"\"{batchName}.txt\"");
            if (exitCode != 0)
            {
                throw new Exception(standardOutput);
            }

            // Convertir les fichiers WMF en JPG
            List<(Uri wmfFileUri, Uri jpegFileUri)> wmfToJpgConversionResults = ConvertBatchWmfToJpg(((dynamic)batch).name);

            // Copie des fichiers .ctt, .pc2 et images JPEG vers serveur
            Uri sourceDirectoryUri = new Uri(CutRiteConfigurationReader.Items["SYSTEM_DATA_PATH"].ToString());
            Uri destinationDirectoryUri = new Uri(new Uri(CutRiteConfigurationReader.Items["MACHINING_CENTER_TRANSFER_PATTERNS_PATH"].ToString()), $"{batchName}/"); 
            Directory.CreateDirectory(destinationDirectoryUri.LocalPath);

            List<(Uri sourceFileUri, Uri destinationFileUri)> filesToCopy = new List<(Uri sourceFileUri, Uri destinationFileUri)>
            {
                (new Uri(sourceDirectoryUri, $"{batchName}.ctt"), new Uri(destinationDirectoryUri, $"{batchName}.ctt")),
                (new Uri(sourceDirectoryUri, $"{batchName}.pc2"), new Uri(destinationDirectoryUri, $"{batchName}.pc2"))
            };
            foreach ((Uri _, Uri jpegFileUri) in wmfToJpgConversionResults)
            {
                filesToCopy.Add((
                    jpegFileUri, 
                    new Uri(destinationDirectoryUri, Path.GetFileName(jpegFileUri.LocalPath))
                ));
            }
            CopyFiles(filesToCopy);

            // Création du fichier de batch "batch.txt"
            File.WriteAllText(new Uri(destinationDirectoryUri, "batch.txt").LocalPath, ((dynamic)batch).id.ToString());

            // Optimisation terminée
            await UpdateEtatMpr(((dynamic)batch).id, 'G');
        }

        /// <summary>
        /// Deletes the part list file of this batch. In doing so, the importation process is simplified somewhat. 
        /// Also, sending an erroneous batch to the machining center is a bit harder.
        /// </summary>
        /// <param name="batchName"> The name of the batch to process. </param>
        private void DeletePartListFile(string batchName)
        {
            string fileName = $"{batchName}.prl";
            string directoryPath = CutRiteConfigurationReader.Items["SYSTEM_PART_LIST_PATH"].ToString();
            string fullFilePath = directoryPath + fileName;

            if (File.Exists(fullFilePath))
            {
                using (FileSystemWatcher fileSystemWatcher = new FileSystemWatcher(directoryPath, fileName))
                using (ManualResetEventSlim threadSynchronizationEvent = new ManualResetEventSlim())
                {
                    fileSystemWatcher.EnableRaisingEvents = true;
                    fileSystemWatcher.Deleted += (object sender, FileSystemEventArgs e) => { threadSynchronizationEvent.Set(); };
                    File.Delete(fullFilePath);

                    int maximumWaitTime = 5;
                    threadSynchronizationEvent.Wait(maximumWaitTime * 1000);
                    if (!threadSynchronizationEvent.IsSet)
                    {
                        throw new Exception($"Deleting part list file \"{fullFilePath}\" took more than {maximumWaitTime} seconds.");
                    }
                }
            }
        }

        /// <summary>
        /// Converts a batch's wmf image files to jpg.
        /// </summary>
        /// <param name="batchName">The batch name.</param>
        /// </param>
        public List<(Uri wmfFileUri, Uri jpegFileUri)> ConvertBatchWmfToJpg(string batchName)
        {
            List<(Uri wmfFileUri, Uri jpegFileUri)>  wmfToJpgConversionResults = new List<(Uri wmfFileUri, Uri jpegFileUri)>();
            for (int i = 1; i < 9999; i++)
            {
                Uri wmfFileUri = new Uri(
                    new Uri(CutRiteConfigurationReader.Items["SYSTEM_DATA_PATH"].ToString()), 
                    "$" + batchName + FillZero(i, 4) + "$.wmf"
                );

                if (File.Exists(wmfFileUri.LocalPath))
                {
                    // Conversion WMF vers JPG
                    Uri jpgFileUri = new Uri(
                        new Uri(CutRiteConfigurationReader.Items["SYSTEM_DATA_PATH"].ToString()), 
                        batchName + FillZero(i, 4) + ".jpg"
                    );
                    wmfToJpgConversionResults.Add((wmfFileUri, jpgFileUri));
                    ExecuteProcess(
                        new Uri(new Uri($"{Application.StartupPath}/"), "autoit/imagemagick.exe").LocalPath, 
                        $"\"{wmfFileUri.LocalPath}\" \"{jpgFileUri.LocalPath}\""
                    );
                }
                else
                {
                    break;
                }
            }

            return wmfToJpgConversionResults;
        }

        public void CopyFiles(List<(Uri sourceFileUri, Uri destinationFileUri)> filesToCopy)
        {
            foreach ((Uri sourceFileUri, Uri destinationFileUri) in filesToCopy)
            {
                File.Copy(sourceFileUri.LocalPath, destinationFileUri.LocalPath, true);
            }
        }

        /// <summary>
        /// Updates the mpr status (and the comments) of the batch
        /// </summary>
        /// <param name="id"> The id of the batch</param>
        /// <param name="mprStatus">The new status of the batch</param>
        /// <param name="comments">The new comments of the batch</param>
        /// <exception cref="Exception">Thrown when a request to Fabplan fails.</exception>
        private async Task UpdateEtatMpr(long id, char mprStatus, string comments = null)
        {
            UriBuilder builder = new UriBuilder()
            {
                Host = ConfigINI.Items["FABPLAN_HOST_NAME"].ToString(),
                Path = ConfigINI.Items["FABPLAN_SET_MPR_STATUS_URL"].ToString(),
                Port = -1,
                Scheme = "http"
            };

            try
            {
                await FabplanHttpRequest.Post(builder.ToString(), new { id, mprStatus });
            }
            catch (Exception e)
            {
                throw new Exception($"Could not assign mpr status \"{mprStatus}\" to batch with unique identifier \"{id}\" in the API.", e);
            }

            if (comments != null)
            {
                builder.Path = ConfigINI.Items["FABPLAN_SET_COMMENTS_URL"].ToString();
                try
                {
                    await FabplanHttpRequest.Post(builder.ToString(), new { id, comments });
                }
                catch (Exception e)
                {
                    throw new Exception($"Could not assign comments \"{comments}\" to batch with unique identifier \"{id}\" in the API.", e);
                }
            }
        }


        /// <summary>
        /// Modifies pannels in the brd file of the current batch
        /// </summary>
        /// <param name="panneaux">The pannel code to use with the current batch</param>
        /// <param name="batchName">The name of the batch</param>
        private void ModifyPanneaux(string panneaux, string batchName)
        {
            string[] pans = panneaux.Split(',');
            Uri boardFileUri = new Uri(new Uri(CutRiteConfigurationReader.Items["SYSTEM_DATA_PATH"].ToString()), $"{batchName}.brd");

            string[] brdLines = File.ReadAllText(boardFileUri.LocalPath).Split(new string[] { "\r\n" }, StringSplitOptions.None);

            string brd = brdLines[0] + "\r\n" + brdLines[1] + "\r\n" + brdLines[2] + "\r\n";

            foreach(string line in brdLines)
            {
                if (line.Length > 5)
                {
                    if (line.Substring(0, 4) == "BRD1")
                    {
                        foreach (string pan in pans)
                        {
                            if (line.IndexOf(pan) > 0)
                            {   
                                // Panneau match
                                brd += line + "\r\nBRD2,,0,1,0,\r\n";
                            }
                        }
                    }
                }
            }

            File.WriteAllText(boardFileUri.LocalPath, brd);
        }

        /// <summary>
        /// Adds 0's at the beginning of the integer <paramref name="i"/> in order to normalize its length to <paramref name="nb"/>.
        /// </summary>
        /// <param name="i"></param>
        /// <param name="nb"></param>
        /// <returns></returns>
        private string FillZero(int i, short nb)
        {
            string ret = i.ToString();
            while (ret.Length < nb)
            {
                ret = "0" + ret;
            }
            return ret;
        }
    }
}