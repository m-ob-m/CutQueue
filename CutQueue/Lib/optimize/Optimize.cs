using CutQueue.Lib.Fabplan;
using CutQueue.Logging;
using Newtonsoft.Json.Linq;
using System;
using System.Dynamic;
using System.IO;
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
                        Logger.Log(new Exception("Could not optimize batch.", e).ToString());
                    }
                    finally
                    {
                        _inProgress = false;
                    }
                }
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
                    await UpdateEtatMpr(batch.id, 'E', e.Message);
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
        private async Task<ExpandoObject> GetNextBatchToOptimize()
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
            if (batch != null)
            {
                try
                {
                    batch.Add("id", rawBatch.GetValue("id"));
                }
                catch (Exception e)
                {
                    throw new Exception("Next batch to optimize is missing its \"id\" meember.", e);
                }

                try
                {
                    batch.Add("name", rawBatch.GetValue("name"));
                }
                catch (Exception e)
                {
                    throw new Exception("Next batch to optimize is missing its \"name\" meember.", e);
                }

                try
                {
                    batch.Add("pannels", rawBatch.GetValue("pannels"));
                }
                catch (Exception e)
                {
                    throw new Exception("Next batch to optimize is missing its \"pannels\" meember.", e);
                }
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
            await UpdateEtatMpr(batch.id, 'P');

            string toImport = batch.name + ".txt";      // Nom du fichier CSV

            // Import du CSV
            ExecuteProcess("autoit\\importCSV.exe", toImport);

            // Génération du fichier des panneaux (.brd)
            ExecuteProcess("autoit\\panneaux.exe ", toImport);

            // Modification du fichier .brd pour avoir les bons panneaux
            ModifyPanneaux(batch.pannels, batch.name);

            // Optimisation
            string msg = ExecuteProcess("autoit\\optimise.exe", toImport);
            if (msg != "OK")
            {	
                // Le script retourne OK seulement quand tout a bien été
                throw new ArgumentException(msg);
            }


            // Convertir les fichiers WMF en JPG
            short nbImg = 0;     // Décompte pour copie des fichiers
            for (int i = 1; i < 9999; i++)
            {
                string wmf_name = "$" + batch.name + FillZero(i, 4) + "$.wmf";

                if (File.Exists(ConfigINI.GetInstance().Items["FABRIDOR"] + "SYSTEM_DATA\\DATA\\" + wmf_name))
                {
                    // Conversion WMF vers JPG
                    ExecuteProcess("autoit\\imagemagick.exe", wmf_name + " " + batch.name + FillZero(i, 4) + ".jpg");	
                    nbImg++;
                }
                else
                {
                    break;
                }
            }

            // Copie des fichiers .ctt, .pc2 et images JPEG vers serveur
            string rep_dest = ConfigINI.GetInstance().Items["V200"] + batch.nameh + "\\"; // Répertoire des destination des fichiers
            Directory.CreateDirectory(rep_dest);

            string sourceCTTFilepath = ConfigINI.GetInstance().Items["FABRIDOR"] + "SYSTEM_DATA\\DATA\\" + batch.name + ".ctt";
            File.Copy(sourceCTTFilepath, rep_dest + batch.name + ".ctt", true);
            string sourcePC2FilePath = ConfigINI.GetInstance().Items["FABRIDOR"] + "SYSTEM_DATA\\DATA\\" + batch.name + ".pc2";
            File.Copy(sourcePC2FilePath, rep_dest + batch.name + ".pc2", true);

            for (short i = 1; i <= nbImg; i++)  // Copie des images JPEG
            {
                string source = ConfigINI.GetInstance().Items["FABRIDOR"] + "SYSTEM_DATA\\DATA\\" + batch.name + FillZero(i, 4) + ".jpg";
                string destination = rep_dest + batch.name + FillZero(i, 4) + ".jpg";
                File.Copy(source, destination, true);
            }

            // Création du fichier de batch "batch.txt"
            File.WriteAllText(rep_dest + "batch.txt", batch.id.ToString());

            // Optimisation terminée
            await UpdateEtatMpr(batch.id, 'G');
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