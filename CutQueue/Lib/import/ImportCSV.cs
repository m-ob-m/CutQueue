using CutQueue.Lib.JobImporter;
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
        private static bool inProgress = false;

        public ImportCSV(){}

        /// <summary>
        /// Importation process
        /// </summary>
        public async Task DoSync()
        {
            if (!inProgress)
            {
                try
                {
                    inProgress = true;
                    int previousUpdateDate = await GetLastUpdateDate();
                    int currentUpdateDate = await ImportAllCSV(previousUpdateDate);
                    if(currentUpdateDate > previousUpdateDate)
                    {
                        await SetLastUpdateDate(currentUpdateDate);
                    }
                }
                catch (Exception e)
                {
                    Logger.Log(string.Format("Error during importation process: {0}", e.ToString()));
                }
                finally
                {
                    inProgress = false;
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
            string userName = ConfigINI.GetInstance().Items["SIA_USER_NAME"].ToString();
            string password = ConfigINI.GetInstance().Items["SIA_PASSWORD"].ToString();
            string domainName = ConfigINI.GetInstance().Items["SIA_DOMAIN_NAME"].ToString();
            bool debug = ConfigINI.GetInstance().Items["DEBUG"].ToString() != "0";
            string debugImportFileName = ConfigINI.GetInstance().Items["DEBUG_IMPORT_FILE_NAME"].ToString();
            DateTime unixEpoch = new DateTime(1970, 1, 1, 0, 0, 0, 0);

            using (new Impersonation(domainName, userName, password))
            {
                FileInfo[] fileInfos = new DirectoryInfo(ConfigINI.GetInstance().Items["CSV"].ToString())
                    .GetFiles()
                    .Where((file) => {
                        int unixTimestamp = Convert.ToInt32(file.LastWriteTimeUtc.Subtract(unixEpoch).TotalSeconds);
                        return !debug && unixTimestamp > highestDate || debug && file.Name == debugImportFileName;
                    })
                    .OrderBy((file) => {return file.LastWriteTimeUtc;})
                    .ToArray();

                foreach (FileInfo fileInfo in fileInfos)
                {
                    await JobImporter.ImportJobFromCSVFileAndExportToFabplan(fileInfo);
                    int unixTimestamp = Convert.ToInt32(fileInfo.LastWriteTimeUtc.Subtract(unixEpoch).TotalSeconds);
                    highestDate = Math.Max(highestDate, unixTimestamp);
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
