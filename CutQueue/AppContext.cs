using CutQueue.Lib.Fabplan;
using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using CutQueue.Logging;


/**
 * \name		AppContext
* \author    	Mathieu Grenier
* \version		1.0
* \date       	2017-02-27
*
* \brief 		Contexte de l'application
* \detail       Contexte de l'application. C'est ici que l'importation des CSV et les optimisations
*               de CutRite sont effectuées (dans leur classe respective)
*/
namespace CutQueue
{
    /// <summary>
    /// The application Context
    /// </summary>
    class AppContext : ApplicationContext
    {
        //Component declarations
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayIconContextMenu;
        private ToolStripMenuItem closeMenuItem;
        private ToolStripMenuItem syncMenuItem;

        private Timer timer;   // Timer de synchronisation
        private readonly ImportCSV importationProcess;
        private readonly Optimize optimizationProcess;

        /// <summary>
        /// Creates the program
        /// </summary>
        public AppContext()
        {
            Application.ApplicationExit += new EventHandler(OnApplicationExit);
            InitializeComponent();
            trayIcon.Visible = true;

            // Synchronization timer
            timer = new Timer
            {
                Interval = int.Parse(ConfigINI.GetInstance().Items["TIMER"].ToString())
            };
            timer.Tick += T_Tick;
            timer.Start();

            // Synchroniztion classes
            importationProcess = new ImportCSV();
            optimizationProcess = new Optimize();
        }

        /// <summary>
        /// Creates the menu items of the application
        /// </summary>
        private void InitializeComponent()
        {
            closeMenuItem = new ToolStripMenuItem
            {
                Name = "CloseMenuItem",
                Size = new Size(152, 22),
                Text = "Fermer CutQueue"
            };
            closeMenuItem.Click += new EventHandler(CloseMenuItem_Click);

            syncMenuItem = new ToolStripMenuItem
            {
                Name = "SyncMenuItem",
                Size = new Size(152, 22),
                Text = "Synchronisation manuelle"
            };
            syncMenuItem.Click += new EventHandler(SyncMenuItem_Click);

            trayIconContextMenu = new ContextMenuStrip
            {
                Name = "TrayIconContextMenu",
                Size = new Size(153, 70)
            };
            trayIconContextMenu.SuspendLayout();
            trayIconContextMenu.Items.AddRange(new ToolStripItem[] {syncMenuItem, closeMenuItem});
            trayIconContextMenu.ResumeLayout(false);

            trayIcon = new NotifyIcon
            {
                BalloonTipIcon = ToolTipIcon.Info,
                BalloonTipText = "Cliquer avec le bouton de droite pour avoir les options.",
                BalloonTipTitle = "CutQueue",
                Text = "CutQueue",
                Icon = Properties.Resources.v9,
                ContextMenuStrip = trayIconContextMenu
            };
            trayIcon.DoubleClick += TrayIcon_DoubleClick;
        }

        /// <summary>
        /// An event synchronized with the synchrnization timer
        /// </summary>
        /// <param name="sender">The element that triggered the event</param>
        /// <param name="e">The arguments of the event</param>
        private void T_Tick(object sender, EventArgs e)
        {
            var temp = DoSync();
        }

       /// <summary>
       /// The synchronization function
       /// </summary>
        private async Task DoSync()
        {
            try
            {
                await LogInToFabplan();
                var temp = Task.Factory.StartNew(async () => await importationProcess.DoSync());
                temp = Task.Factory.StartNew(async () => await optimizationProcess.DoOptimize());
            }
            catch (Exception e)
            {
                Logger.Log(e.ToString());
            }
        }

        /// <summary>
        /// A function that is called upon exiting the application.
        /// </summary>
        /// <param name="sender">The element that triggered the event</param>
        /// <param name="e">The arguments of the event</param>
        private async void OnApplicationExit(object sender, EventArgs eventArguments)
        {
            try
            {
                await LogOutFromFabplan();
            }
            catch (Exception e)
            {
                Logger.Log(e.ToString());
            }

            //Cleanup so that the icon will be removed when the application is closed
            trayIcon.Visible = false;
        }

        /// <summary>
        /// A function that is called upon double clicking the tray icon.
        /// </summary>
        /// <param name="sender">The element that triggered the event</param>
        /// <param name="e">The arguments of the event</param>
        private void TrayIcon_DoubleClick(object sender, EventArgs e)
        {
            //Here you can do stuff if the tray icon is doubleclicked
            trayIcon.ShowBalloonTip(10000);
        }

        /// <summary>
        /// A function that is called upon clicking the synchronize menu entry in the tray icon menu.
        /// </summary>
        /// <param name="sender">The element that triggered the event</param>
        /// <param name="e">The arguments of the event</param>
        private void SyncMenuItem_Click(object sender, EventArgs e)
        {
            var temp = DoSync();
        }

        /// <summary>
        /// A function that is called upon clicking the close menu entry in the tray icon menu.
        /// </summary>
        /// <param name="sender">The element that triggered the event</param>
        /// <param name="e">The arguments of the event</param>
        private void CloseMenuItem_Click(object sender, EventArgs e)
        {
            string message = "Voulez-vous vraiment fermer CutQueue?";
            string title = "Confirmation";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            MessageBoxIcon icon = MessageBoxIcon.Exclamation;
            MessageBoxDefaultButton button = MessageBoxDefaultButton.Button2;
            DialogResult result = DialogResult.Yes;
            if (MessageBox.Show(message, title, buttons, icon, button) == result)
            {
                Application.Exit();
            }
        }

        private async Task LogInToFabplan()
        {
            UriBuilder builder = new UriBuilder()
            {
                Host = ConfigINI.GetInstance().Items["HOST_NAME"].ToString(),
                Path = ConfigINI.GetInstance().Items["LOGIN_URL"].ToString(),
                Port = -1,
                Scheme = "http"
            };

            dynamic credentials = new {
                username = ConfigINI.GetInstance().Items["FABPLAN_USER_NAME"],
                password = ConfigINI.GetInstance().Items["FABPLAN_PASSWORD"]
            };

            try
            {
                await FabplanHttpRequest.Post(builder.ToString(), credentials);
            }
            catch (FabplanHttpResponseWarningException e)
            {
                throw new Exception("Couldn't log in to fabplan as user \"" + credentials.username + "\".", e);
            }
        }

        private async Task LogOutFromFabplan()
        {
            UriBuilder builder = new UriBuilder()
            {
                Host = ConfigINI.GetInstance().Items["HOST_NAME"].ToString(),
                Path = ConfigINI.GetInstance().Items["LOGOUT_URL"].ToString(),
                Port = -1,
                Scheme = "http"
            };

            try
            {
                await FabplanHttpRequest.Post(builder.ToString(), new { });
            }
            catch (FabplanHttpResponseWarningException e)
            {
                throw new Exception("Couldn't log out from fabplan.", e);
            }
        }
    }
}
