﻿using CutQueue.Lib.import.model;
using System;
using System.IO;
using System.Linq;
using CutQueue.Lib.tools;
using CutQueue.Lib.Fabplan;
using System.Threading.Tasks;
using CutQueue.Logging;

/**
 * \name		ImportCSV
* \author    	Mathieu Grenier
* \version		1.0
* \date       	2017-02-27
*
* \brief 		Vérifie s'il y a de nouveaux fichiers CSV à importer
* \detail       Vérifie s'il y a de nouveaux fichiers CSV à importer
*/
namespace CutQueue
{
    class ImportCSV
    {
        private static bool _inProgress = false;

        public ImportCSV()
        {
        }

        /// <summary>
        /// Importation process
        /// </summary>
        public async Task DoSync()
        {
            if (!_inProgress)
            {
                try
                {
                    _inProgress = true;
                    int previousUpdateDate = await GetLastUpdateDate();
                    int currentUpdateDate = await ImportAllCSV(previousUpdateDate);
                    if(currentUpdateDate > previousUpdateDate)
                    {
                        await SetLastUpdateDate(currentUpdateDate);
                    }
                }
                catch (Exception e)
                {
                    throw new Exception("Error during importation process.", e);
                }
                finally
                {
                    _inProgress = false;
                }
            }
        }



        /// <summary>
        /// Imports all csv files that were created since the last update date of the importator module in Fabplan.
        /// </summary>
        /// <param name="highestDate">The UNIX timestamp of the last update date of the importator module in Fabplan</param>
        /// <returns>The UNIX timestamp of the last update date of the importator module in Fabplan</returns>
        private async Task<int> ImportAllCSV(int highestDate)
        {
            string userName = ConfigINI.GetInstance().Items["SIA_username"].ToString();
            string password = ConfigINI.GetInstance().Items["SIA_password"].ToString();
            using (new Impersonation(userName, "", password))
            {
                FileInfo[] files = new DirectoryInfo(ConfigINI.GetInstance().Items["CSV"].ToString())
                    .GetFiles()
                    .Where(p => (int)p.LastWriteTimeUtc.Subtract(new DateTime(1970, 1, 1, 0, 0, 0, 0)).TotalSeconds > highestDate)
                    .OrderBy(p => p.LastWriteTimeUtc).ToArray();

                foreach (FileInfo file in files)
                {
                    await new FichierCSV(file).Import();
                    highestDate = (int)file.LastWriteTimeUtc.Subtract(new DateTime(1970, 1, 1, 0, 0, 0, 0)).TotalSeconds;
                }
            }

            return highestDate;
        }

        /// <summary>
        /// Gets the last update date of the importator module in Fabplan
        /// </summary>
        /// <exception cref="Exception">Thrown when The request to fabplan fails</exception>
        /// <returns>The unix timestamp of the last update date</returns>
        private async Task<int> GetLastUpdateDate()
        {
            UriBuilder builder = new UriBuilder()
            {
                Host = ConfigINI.GetInstance().Items["HOST_NAME"].ToString(),
                Path = ConfigINI.GetInstance().Items["GET_LAST_UPDATE_DATE_URL"].ToString(),
                Port = -1,
                Scheme = "http"
            };

            string rawResponse = null;
            try
            {
                rawResponse = await FabplanHttpRequest.Get(builder.ToString());
            }
            catch (Exception e)
            {
                throw new Exception("Last update timestamp could not be retrieved from the API.", e);
            }

            try
            {
                return int.Parse(rawResponse);
            }
            catch (Exception e)
            {
                throw new Exception("Last update timestamp retrieved from the API cannot be converted to a UNIX timestamp.", e);
            }
        }

        /// <summary>
        /// Sets the last update date of the importator module in Fabplan
        /// </summary>
        /// <param name="lastUpdateDate">A UNIX timestamp that represents the last update date</param>
        /// <exception cref="Exception">Thrown when The request to fabplan fails</exception>
        private async Task SetLastUpdateDate(int lastUpdateDate)
        {
            UriBuilder builder = new UriBuilder()
            {
                Host = ConfigINI.GetInstance().Items["HOST_NAME"].ToString(),
                Path = ConfigINI.GetInstance().Items["SET_LAST_UPDATE_DATE_URL"].ToString(),
                Port = -1,
                Scheme = "http"
            };

            try
            {
                await FabplanHttpRequest.Post(builder.ToString(), new { lastUpdateDate });
            }
            catch (Exception e)
            {
                throw new Exception("Last update timestamp could not be set in the API.", e);
            }
        }
    }
}
