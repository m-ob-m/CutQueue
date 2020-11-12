using System.Collections;
using System.IO;
using System.Windows.Forms;
using static System.Environment;

namespace CutQueue.Lib
{
    class CutRiteConfigurationReader
    {
        private static Hashtable items = null;

        /// <summary>
        /// Gets Cut Rite configuration from configuration files.
        /// </summary>
        private static void GetItems()
        {
            items = new Hashtable();

            string configurationFilePath = Path.Combine(
                ConfigINI.Items["CUT_QUEUE_INSTALLATION_PATH"].ToString(),
                ConfigINI.Items["CUT_QUEUE_WORKSPACE"].ToString(),
                "systemv9.ctl"
            );

            if (File.Exists(configurationFilePath))
            {
                ReadSystemV9CtlConfigurationFile(configurationFilePath);
            }
            else
            {
                MessageBox.Show($"Cut Queue configuration file \"{configurationFilePath}\" not found.");
                Exit(0);
            }

            configurationFilePath = Path.Combine(
                ConfigINI.Items["CUT_QUEUE_INSTALLATION_PATH"].ToString(),
                ConfigINI.Items["CUT_QUEUE_WORKSPACE"].ToString(),
                "mmch.ctl"
            );

            if (File.Exists(configurationFilePath))
            {
                ReadMmchCtlConfigurationFile(configurationFilePath);
            }
            else
            {
                MessageBox.Show($"Cut Queue configuration file \"{configurationFilePath}\" not found.");
                Exit(0);
            }
        }

        /// <summary>
        /// Reads a CutRite's "systemv9.ctl" configuration file.
        /// </summary>
        /// <param name="configurationFilePath">The path of the CutRite systemv9.ctl file to read</param>
        private static void ReadSystemV9CtlConfigurationFile(string configurationFilePath)
        {
            using (StreamReader streamReader = new StreamReader(configurationFilePath))
            {
                while (streamReader.Peek() != -1)
                {
                    string[] parameter = streamReader.ReadLine().Split(',');

                    if (parameter.Length == 2)
                    {
                        if (parameter[0] == "SYSPARTLISTPATH")
                        {
                            items.Add("SYSTEM_PART_LIST_PATH", parameter[1]);
                        }
                        else if (parameter[0] == "SYSDATAPATH")
                        {
                            items.Add("SYSTEM_DATA_PATH", parameter[1]);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Reads a CutQueue's "mmch.ctl" configuration file.
        /// </summary>
        /// <param name="configurationFilePath">The path of the CutQueue mmch.ctl file to read</param>
        private static void ReadMmchCtlConfigurationFile(string configurationFilePath)
        {
            using (StreamReader streamReader = new StreamReader(configurationFilePath))
            {
                while (streamReader.Peek() != -1)
                {
                    string[] parameter = streamReader.ReadLine().Split(',');

                    if (parameter.Length == 13)
                    {
                        if (!items.ContainsKey("MACHINING_CENTER_TRANSFER_PATTERNS_PATH") && parameter[0] == "MCH_TRANSINFO")
                        {
                            items.Add("MACHINING_CENTER_TRANSFER_PATTERNS_PATH", parameter[12]);
                        }
                    }
                }
            }
        }

        public static Hashtable Items
        {
            get
            {
                if (items == null)
                {
                    GetItems();
                }

                return items;
            }
        }
    }
}
