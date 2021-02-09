using CutQueue.Lib;
using CutQueue.Lib.Fabplan;
using CutQueue.Lib.Tools;
using CutQueue.Logging;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CutQueue
{
    /// <summary>
    /// A class that optimizes batches from Fabplan using CutRite to create machining programs
    /// </summary>
    class Optimize
    {
        private static bool _inProgress = false;
        private enum ContinueStatus { No, Yes }

        /// <summary>
        /// Main constructor
        /// </summary>
        public Optimize() { }

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
                        Logger.Log(e.ToString());
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
            dynamic batch;
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
                    TransferToMachiningCenter(batch);
                    await UpdateEtatMpr(((dynamic)batch).id, 'G');
                }
                catch (Exception e)
                {
                    Logger.Log(e.ToString() + "\n");
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
        /// Optimizes a given batch using Cut Rite.
        /// </summary>
        /// <param name="batch"> The name of the batch to optimize</param>
        /// <exception cref="ArgumentException">Thrown when optimise.exe fails.</exception>
        private async Task OptimizeBatch(dynamic batch)
        {
            // Changer l'état pour en cours (P pour In Progress)
            await UpdateEtatMpr(((dynamic)batch).id, 'C');

            string batchName = ((dynamic)batch).name.ToString();      // Nom du fichier CSV

            DeleteCutRiteBatchData(batchName);

            int exitCode; string standardOutput; string standardError;

            // Import du CSV
            (exitCode, standardOutput, standardError) = ProcessAdvanced.ExecuteProcess(
                Path.Combine(Application.StartupPath, "AutoIt/ImportCSV.exe"),
                $"\"{batchName}.txt\"",
                int.Parse(ConfigINI.Items["MAXIMUM_EXECUTION_TIME_FOR_IMPORT_EXE"].ToString())
            );
            if (exitCode != 0)
            {
                throw new Exception(standardOutput + standardError);
            }

            // Génération du fichier des panneaux (.brd)
            (exitCode, standardOutput, standardError) = ProcessAdvanced.ExecuteProcess(
                Path.Combine(Application.StartupPath, "Autoit/Panneaux.exe"),
                $"\"{batchName}.txt\"",
                int.Parse(ConfigINI.Items["MAXIMUM_EXECUTION_TIME_FOR_PANNEAUX_EXE"].ToString())
            );
            if (exitCode != 0)
            {
                throw new Exception(standardOutput + standardError);
            }

            // Modification du fichier .brd pour avoir les bons panneaux
            ModifyPanneaux(((dynamic)batch).pannels, ((dynamic)batch).name);

            // Optimisation
            (exitCode, standardOutput, standardError) = ProcessAdvanced.ExecuteProcess(
                Path.Combine(Application.StartupPath, "AutoIt/Optimise.exe"),
                $"\"{batchName}.txt\"",
                int.Parse(ConfigINI.Items["MAXIMUM_EXECUTION_TIME_FOR_OPTIMISE_EXE"].ToString())
            );
            if (exitCode != 0)
            {
                throw new Exception(standardOutput + standardError);
            }
        }

        /// <summary>
        /// Transfers a batch to the machining center.
        /// </summary>
        /// <param name="batch">The name of the batch</param>
        public void TransferToMachiningCenter(dynamic batch)
        {
            string batchName = ((dynamic)batch).name.ToString();

            // Convertir les fichiers WMF en JPG
            List<(string wmfFile, string jpgFile)> wmfToJpgConversionResults = ConvertBatchWmfToJpg(batchName);

            // Copie des fichiers .ctt, .pc2 et images JPEG vers serveur
            string sourceDirectory = CutRiteConfigurationReader.Items["SYSTEM_DATA_PATH"].ToString();
            string destinationDirectory = Path.Combine(CutRiteConfigurationReader.Items["MACHINING_CENTER_TRANSFER_PATTERNS_PATH"].ToString(), batchName);

                
            FileAdvanced.Delete(destinationDirectory);
            Directory.CreateDirectory(destinationDirectory);

            // Simplifier les programmes d'usinage.
            SimplifyMprFiles(batchName);

            // Copier les fichiers pertinents au centre d'usinage.
            File.Copy(Path.Combine(sourceDirectory, $"{batchName}.ctt"), Path.Combine(destinationDirectory, $"{batchName}.ctt"), true);
            File.Copy(Path.Combine(sourceDirectory, $"{batchName}.pc2"), Path.Combine(destinationDirectory, $"{batchName}.pc2"), true);
            wmfToJpgConversionResults.ForEach(result => File.Copy(result.jpgFile, Path.Combine(destinationDirectory, Path.GetFileName(result.jpgFile)), true));

            // Création du fichier de batch "batch.txt"
            File.WriteAllText(Path.Combine(destinationDirectory, "batch.txt"), ((dynamic)batch).id.ToString());
        }

        /// <summary>
        /// Simplifies the mpr files for a batch.
        /// </summary>
        /// <param name="batchName">The name of the batch</param>
        private void SimplifyMprFiles(string batchName)
        {
            foreach (string sourceMprFile in GetNestingMprFiles(batchName))
            {
                string destinationMprFile = Path.Combine(
                    CutRiteConfigurationReader.Items["MACHINING_CENTER_TRANSFER_PATTERNS_PATH"].ToString(),
                    batchName,
                    $"{Path.GetFileNameWithoutExtension(sourceMprFile)}.mpr".Replace("$", "")
                );

                (int exitCode, string standardOutput, string standardError) = ProcessAdvanced.ExecuteProcess(
                    Path.Combine(Application.StartupPath, "MprSimplifier.exe"),
                    $"\"--input-file\" \"{sourceMprFile}\" \"--output-file\" \"{destinationMprFile}\"",
                    int.Parse(ConfigINI.Items["MAXIMUM_EXECUTION_TIME_FOR_MPRSIMPLIFIER_EXE"].ToString())
                );

                if (exitCode != 0)
                {
                    throw new Exception(standardOutput + standardError);
                }
            }
        }

        /// <summary>
        /// Deletes the part list file of this batch. In doing so, the importation process is simplified somewhat. 
        /// Also, sending an erroneous batch to the machining center is a bit harder.
        /// </summary>
        /// <param name="batchName"> The name of the batch to process. </param>
        private void DeleteCutRiteBatchData(string batchName)
        {
            List<string> filesToDelete = GetNestingMprFiles(batchName);
            filesToDelete.AddRange(GetNestingWmfFiles(batchName));
            filesToDelete.Add(Path.Combine(CutRiteConfigurationReader.Items["SYSTEM_PART_LIST_PATH"].ToString(), $"{batchName}.prl"));
            filesToDelete.ForEach(fileToDelete => File.Delete(fileToDelete));
        }

        /// <summary>
        /// Returns the list of nesting mpr files for a given batch. 
        /// </summary>
        /// <param name="batchName">The batch name.</param>
        private List<string> GetNestingMprFiles(string batchName)
        {
            return Directory.GetFiles(CutRiteConfigurationReader.Items["SYSTEM_PART_LIST_PATH"].ToString())
                .Where(path => new Regex(@"(?<=[\\/])\$" + Regex.Escape(batchName) + @"\d{4}\$\.mpr\z", RegexOptions.IgnoreCase).IsMatch(path))
                .ToList();
        }

        /// <summary>
        /// Returns the list of nesting wmf files for a given batch. 
        /// </summary>
        /// <param name="batchName">The batch name.</param>
        private List<string> GetNestingWmfFiles(string batchName)
        {
            return Directory.GetFiles(CutRiteConfigurationReader.Items["SYSTEM_PART_LIST_PATH"].ToString())
                .Where(path => new Regex(@"(?<=[\\/])\$" + Regex.Escape(batchName) + @"\d{4}\$\.wmf\z", RegexOptions.IgnoreCase).IsMatch(path))
                .ToList();
        }

        /// <summary>
        /// Converts a batch's wmf image files to jpg.
        /// </summary>
        /// <param name="batchName">The batch name.</param>
        private List<(string wmfFile, string jpegFile)> ConvertBatchWmfToJpg(string batchName)
        {
            List<(string wmfFile, string jpegFile)> wmfToJpgConversionResults = new List<(string wmfFile, string jpegFile)>();
            foreach (string wmfFile in GetNestingWmfFiles(batchName))
            {
                // Conversion WMF vers JPG
                string jpgFile = Path.Combine(Path.GetDirectoryName(wmfFile), $"{Path.GetFileNameWithoutExtension(wmfFile)}.jpg".Replace("$", ""));
                wmfToJpgConversionResults.Add((wmfFile, jpgFile));
                ProcessAdvanced.ExecuteProcess(
                    Path.Combine(Application.StartupPath, "AutoIt/ImageMagick.exe"),
                    $"\"{wmfFile}\" \"{jpgFile}\"",
                    int.Parse(ConfigINI.Items["MAXIMUM_EXECUTION_TIME_FOR_IMAGEMAGIK_EXE"].ToString())
                );
            }

            return wmfToJpgConversionResults;
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
            string boardFile = Path.Combine(CutRiteConfigurationReader.Items["SYSTEM_DATA_PATH"].ToString(), $"{batchName}.brd");

            string[] brdLines = File.ReadAllText(boardFile).Split(new string[] { "\r\n" }, StringSplitOptions.None);

            string brd = brdLines[0] + "\r\n" + brdLines[1] + "\r\n" + brdLines[2] + "\r\n";

            foreach (string line in brdLines)
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

            File.WriteAllText(boardFile, brd);
        }
    }
}