using CutQueue.Lib.Fabplan;
using CutQueue.Logging;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

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
                Host = ConfigINI.GetInstance().Items["HOST_NAME"].ToString(),
                Path = ConfigINI.GetInstance().Items["GET_NEXT_BATCH_TO_OPTIMIZE_URL"].ToString(),
                Port = -1,
                Scheme = "http"
            };

            JObject rawBatch = null;
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
        private string ExecuteProcess(string procName, string arguments)
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

            string msg = process.StandardOutput.ReadToEnd();

            process.WaitForExit();

            return msg;
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

            string toImport = ((dynamic)batch).name + ".txt";      // Nom du fichier CSV

            // Import du CSV
            ExecuteProcess("autoit\\importCSV.exe", toImport);

            // Génération du fichier des panneaux (.brd)
            ExecuteProcess("autoit\\panneaux.exe ", toImport);

            // Modification du fichier .brd pour avoir les bons panneaux
            ModifyPanneaux(((dynamic)batch).pannels, ((dynamic)batch).name);

            // Optimisation
            string msg = ExecuteProcess("autoit\\optimise.exe", toImport);
            if (msg != "OK")
            {	
                // Le script retourne OK seulement quand tout a bien été
                throw new ArgumentException(msg);
            }

            // Convertir les fichiers WMF en JPG
            ConvertBatchWmfToJpg(((dynamic)batch).name, out List<Tuple<Uri, Uri>> wmfToJpgConversionResults);

            // Copie des fichiers .ctt, .pc2 et images JPEG vers serveur
            Uri sourceDirectoryUri = new Uri(new Uri(new Uri(ConfigINI.GetInstance().Items["FABRIDOR"].ToString()), "SYSTEM_DATA\\"), "DATA\\");
            Uri destinationDirectoryUri = new Uri(new Uri(ConfigINI.GetInstance().Items["V200"].ToString()), (string)((dynamic)batch).name + "\\"); 
            Directory.CreateDirectory(destinationDirectoryUri.LocalPath);

            List<Tuple<Uri, Uri>> filesToCopy = new List<Tuple<Uri, Uri>>
            {
                new Tuple<Uri, Uri>(new Uri(sourceDirectoryUri, ((dynamic)batch).name + ".ctt"), new Uri(destinationDirectoryUri, ((dynamic)batch).name + ".ctt")),
                new Tuple<Uri, Uri>(new Uri(sourceDirectoryUri, ((dynamic)batch).name + ".pc2"), new Uri(destinationDirectoryUri, ((dynamic)batch).name + ".pc2"))
            };
            foreach (Tuple<Uri, Uri> wmfToJpgConversionResult in wmfToJpgConversionResults)
            {
                filesToCopy.Add(
                    new Tuple<Uri, Uri>(
                        wmfToJpgConversionResult.Item2, 
                        new Uri(destinationDirectoryUri, Path.GetFileName(wmfToJpgConversionResult.Item2.ToString()))
                    )
                );
            }
            CopyFiles(filesToCopy);

            // Création du fichier de batch "batch.txt"
            File.WriteAllText(new Uri(destinationDirectoryUri, "batch.txt").LocalPath, ((dynamic)batch).id.ToString());

            // Optimisation terminée
            await UpdateEtatMpr(((dynamic)batch).id, 'G');
        }

        /// <summary>
        /// Converts a batch's wmf image files to jpg.
        /// </summary>
        /// <param name="batchName">The batch name.</param>
        /// <param name="wmfToJpgConversionResults">
        /// A list returned in the following form [wmfFileUri_1 => jpgFileUri_1, wmfFileUri_2 => jpgFileUri_2, ..., wmfFileUri_n => jpgFileUri_n].
        /// </param>
        public void ConvertBatchWmfToJpg(string batchName, out List<Tuple<Uri, Uri>> wmfToJpgConversionResults)
        {
            wmfToJpgConversionResults = new List<Tuple<Uri, Uri>>();
            Uri baseUri = new Uri(ConfigINI.GetInstance().Items["FABRIDOR"] + "SYSTEM_DATA\\DATA\\");
            for (int i = 1; i < 9999; i++)
            {
                Uri wmfFileUri = new Uri(baseUri, "$" + batchName + FillZero(i, 4) + "$.wmf");

                if (File.Exists(wmfFileUri.LocalPath))
                {
                    // Conversion WMF vers JPG
                    Uri jpgFileUri = new Uri(baseUri, batchName + FillZero(i, 4) + ".jpg");
                    wmfToJpgConversionResults.Add(new Tuple<Uri, Uri>(wmfFileUri, jpgFileUri));
                    ExecuteProcess(
                        "autoit\\imagemagick.exe", 
                        Path.GetFileName(wmfFileUri.LocalPath) + " " + Path.GetFileName(jpgFileUri.LocalPath)
                    );
                }
                else
                {
                    break;
                }
            }
        }

        public void CopyFiles(List<Tuple<Uri, Uri>> filesToCopy)
        {
            foreach (Tuple<Uri, Uri> fileToCopy in filesToCopy)
            {
                File.Copy(fileToCopy.Item1.LocalPath, fileToCopy.Item2.LocalPath, true);
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
                Host = ConfigINI.GetInstance().Items["HOST_NAME"].ToString(),
                Path = ConfigINI.GetInstance().Items["SET_MPR_STATUS_URL"].ToString(),
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
                builder.Path = ConfigINI.GetInstance().Items["SET_COMMENTS_URL"].ToString();
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
        /// <param name="nom_batch">The name of the batch</param>
        private void ModifyPanneaux(string panneaux, string nom_batch)
        {
            string[] pans = panneaux.Split(',');
            string brdPath = ConfigINI.GetInstance().Items["FABRIDOR"] + "\\SYSTEM_DATA\\DATA\\" + nom_batch + ".brd";

            string[] brdLines = File.ReadAllText(brdPath).Split(new string[] { "\r\n" }, StringSplitOptions.None);

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

            File.WriteAllText(brdPath, brd);
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