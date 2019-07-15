using CutQueue.Lib.SingleGlobalInstance;
using System;
using System.Windows.Forms;

/**
 * \name		Program
* \author    	Mathieu Grenier
* \version		1.0
* \date       	2017-02-27
*
* \brief 		Point d'entré exécuté lors du démarrage de l'application
*/


namespace CutQueue
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            using (new SingleGlobalInstance(1000)) //1000ms timeout on global lock
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new AppContext());
            }
        }
    }
}
