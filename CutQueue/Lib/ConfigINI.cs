using System;
using System.Collections;
using System.IO;
using System.Windows.Forms;
using static System.Environment;

/**
 * \name		ConfigINI
* \author    	Mathieu Grenier
* \version		1.0
* \date       	2017-02-27
*
* \brief 		Cette classe Singleton lit le fichier INI et organise les lignes de DB du fichier config.ini
*/
namespace CutQueue
{
    class ConfigINI
    {
        private static Hashtable items = null;

        /// <summary>
        /// Gets static instance of class.
        /// </summary>
        /// <returns>The static instance of the class</returns>
        private static void GetItems()
        {
            items = new Hashtable();

            string configurationFilePath = new Uri(new Uri($"{Application.StartupPath}/"), "config.ini").LocalPath;

            if (File.Exists(configurationFilePath))
            {
                ReadConfigurationFile(configurationFilePath);
            }
            else
            {
                MessageBox.Show($"Cut Queue configuration file \"{configurationFilePath}\" not found.");
                Exit(0);
            }

        }
        /// <summary>
        /// Reads an ini file and extracts its properties
        /// </summary>
        /// <param name="configurationFilePath">The filepath of the ini file to read</param>
        private static void ReadConfigurationFile(string configurationFilePath)
        {
            using (StreamReader streamReader = new StreamReader(configurationFilePath))
            {
                while (streamReader.Peek() != -1)
                {
                    string line = streamReader.ReadLine();
                    if (line.IndexOf('=') > 0)
                    {
                        string[] value = line.Split('=');
                        items.Add(value[0], value[1]);
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
