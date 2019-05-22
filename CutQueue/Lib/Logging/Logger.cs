namespace CutQueue.Logging
{
    using System.IO;
    using System;

    /// <summary>
    /// Class that handles logging.
    /// </summary>
    public static class Logger
    {
        private static string text = "";

        /// <summary>
        /// Logs <paramref name="text"/>  in the logfile.
        /// </summary>
        /// <param name="text"></param>
        public static void Log(string text)
        {
            Uri applicationDirectoryUri = new Uri(AppDomain.CurrentDomain.BaseDirectory);
            Uri applicationLogFileUri = new Uri(applicationDirectoryUri, "log.txt");
            string applicationLogFilePath = Uri.UnescapeDataString(applicationLogFileUri.LocalPath);
            string currentDateTime = DateTime.Now.ToString("F");
            Logger.text += $"{currentDateTime}\t{text}\n";
            try
            {
                using (StreamWriter logFile = new StreamWriter(applicationLogFilePath, true))
                {
                    logFile.WriteLine(Logger.text);
                    logFile.Close();
                }
                Logger.text = "";
            }
            catch (Exception)
            {
                Logger.text += $"The previous error was logged at {currentDateTime} due to log file unavailability.\n" ;
            }
        }
    }
}