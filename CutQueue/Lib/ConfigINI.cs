using System.Collections;
using static System.Environment;
using System.IO;
using System.Windows.Forms;

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
        private Hashtable _items = new Hashtable();
        private static ConfigINI _instance = null;

        /// <summary>
        /// Gets static instance of class.
        /// </summary>
        /// <returns>The static instance of the class</returns>
        public static ConfigINI GetInstance()
        {
            if (_instance == null)
            {
                _instance = new ConfigINI(Application.StartupPath + "\\config.ini");
            }

            return _instance;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="sINIPath">The filepath of the ini file to read</param>
        private ConfigINI(string sINIPath)
        {

            if (File.Exists(sINIPath))
            {
                ReadINIFile(sINIPath);
            }
            else
            {
                MessageBox.Show("File config.ini not found!");
                Exit(0);
            }
        }

        /// <summary>
        /// Reads an ini file and extracts its properties
        /// </summary>
        /// <param name="sINIPath">The filepath of the ini file to read</param>
        private void ReadINIFile(string sINIPath)
        {
            StreamReader objReader = new StreamReader(sINIPath);
            string str_line = null;
            string[] str_value = null;

            while (objReader.Peek() != -1)
            {
                str_line = objReader.ReadLine();
                if(str_line.IndexOf('=') > 0)
                {
                    str_value = str_line.Split('=');
                   Items.Add(str_value[0], str_value[1]);
                }
            }
        }

        public Hashtable Items
        {
            get
            {
                return _items;
            }

            set
            {
                _items = value;
            }
        }
    }
}
