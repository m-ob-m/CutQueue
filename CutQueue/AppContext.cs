using CutQueue.Lib.Fabplan;
using CutQueue.Logging;
using Microsoft.VisualBasic;
using System;
using System.Drawing;
using System.Dynamic;
using System.Threading.Tasks;
using System.Windows.Forms;

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
        private NotifyIcon TrayIcon;
        private ContextMenuStrip TrayIconContextMenu;
        private ToolStripMenuItem CloseMenuItem;
        private ToolStripMenuItem SyncMenuItem;
        private ToolStripMenuItem TransferToMachiningCenterMenuItem;
        private readonly Timer timer;   // Timer de synchronisation
        private readonly ImportCSV csv;
        private readonly Optimize optimize;

        /// <summary>
        /// Creates the program
        /// </summary>
        public AppContext()
        {
            Application.ApplicationExit += new EventHandler(OnApplicationExit);
            InitializeComponent();
            TrayIcon.Visible = true;

            // Synchronization timer
            timer = new Timer
            {
                Interval = int.Parse(ConfigINI.Items["TIMER"].ToString())
            };
            timer.Tick += Tick;
            timer.Start();

            // Synchroniztion classes
            csv = new ImportCSV();
            optimize = new Optimize();
        }

        /// <summary>
        /// Creates the menu items of the application
        /// </summary>
        private void InitializeComponent()
        {
            CloseMenuItem = new ToolStripMenuItem
            {
                Name = "CloseMenuItem",
                Size = new Size(152, 22),
                Text = "Fermer CutQueue"
            };
            CloseMenuItem.Click += new EventHandler(CloseMenuItem_Click);

            SyncMenuItem = new ToolStripMenuItem
            {
                Name = "SyncMenuItem",
                Size = new Size(152, 22),
                Text = "Synchronisation manuelle"
            };
            SyncMenuItem.Click += new EventHandler(SyncMenuItem_Click);

            TransferToMachiningCenterMenuItem = new ToolStripMenuItem
            {
                Name = "TransferToMachiningCenterMenuItem",
                Size = new Size(152, 22),
                Text = "Envoyer vers le centre d'usinage",
            };
            TransferToMachiningCenterMenuItem.Click += new EventHandler(TransferToMachiningCenterMenuItem_Click);

            TrayIconContextMenu = new ContextMenuStrip
            {
                Name = "TrayIconContextMenu",
                Size = new Size(153, 70)
            };
            TrayIconContextMenu.SuspendLayout();
            TrayIconContextMenu.Items.AddRange(new ToolStripItem[] { SyncMenuItem, TransferToMachiningCenterMenuItem, CloseMenuItem });
            TrayIconContextMenu.ResumeLayout(false);

            TrayIcon = new NotifyIcon
            {
                BalloonTipIcon = ToolTipIcon.Info,
                BalloonTipText = "Cliquer avec le bouton de droite pour avoir les options.",
                BalloonTipTitle = "CutQueue",
                Text = "CutQueue",
                Icon = Properties.Resources.v9,
                ContextMenuStrip = TrayIconContextMenu
            };
            TrayIcon.DoubleClick += TrayIcon_DoubleClick;
        }

        /// <summary>
        /// An event synchronized with the synchrnization timer
        /// </summary>
        /// <param name="sender">The element that triggered the event</param>
        /// <param name="e">The arguments of the event</param>
        private void Tick(object sender, EventArgs e)
        {
            _ = DoSync();
        }

        /// <summary>
        /// The function that updates images for a given batch.
        /// </summary>
        /// <param name="sender">The element that triggered the event</param>
        /// <param name="e">The arguments of the event</param>
        private async void TransferToMachiningCenterMenuItem_Click(object sender, EventArgs eventArguments)
        {
            try
            {
                dynamic batch = new ExpandoObject();
                batch.name = Interaction.InputBox(
                    "Entrer le nom de la batch à envoyer au centre d'usinage.",
                    "Envoi d'une batch au centre d'usinage",
                    null
                );

                UriBuilder builder = new UriBuilder()
                {
                    Host = ConfigINI.Items["FABPLAN_HOST_NAME"].ToString(),
                    Path = ConfigINI.Items["FABPLAN_GET_BATCH_ID_BY_BATCH_NAME"].ToString(),
                    Port = -1,
                    Scheme = "http"
                };

                try
                {
                    await LogInToFabplan();
                    batch.id = await FabplanHttpRequest.Get(builder.ToString(), new { batch.name });
                }
                catch (Exception e)
                {
                    throw new Exception($"Could not obtain unique identifier from name \"{batch.name}\" in the API.", e);
                }

                optimize.TransferToMachiningCenter(batch);
            }
            catch (Exception e)
            {
                Logger.Log(e.ToString() + "\n");
            }
        }

        /// <summary>
        /// The synchronization function
        /// </summary>
        private async Task DoSync()
        {
            await LogInToFabplan();
            _ = Task.Factory.StartNew(async () => await csv.DoSync());
            _ = Task.Factory.StartNew(async () => await optimize.DoOptimize());
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
                Logger.Log(e.ToString() + "\n");
            }

            //Cleanup so that the icon will be removed when the application is closed
            TrayIcon.Visible = false;
        }

        /// <summary>
        /// A function that is called upon double clicking the tray icon.
        /// </summary>
        /// <param name="sender">The element that triggered the event</param>
        /// <param name="e">The arguments of the event</param>
        private void TrayIcon_DoubleClick(object sender, EventArgs e)
        {
            //Here you can do stuff if the tray icon is doubleclicked
            TrayIcon.ShowBalloonTip(10000);
        }

        /// <summary>
        /// A function that is called upon clicking the synchronize menu entry in the tray icon menu.
        /// </summary>
        /// <param name="sender">The element that triggered the event</param>
        /// <param name="e">The arguments of the event</param>
        private async void SyncMenuItem_Click(object sender, EventArgs e)
        {
            await DoSync();
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

        /// <summary>
        /// Logs in to Fabplan.
        /// </summary>
        private async Task LogInToFabplan()
        {
            UriBuilder builder = new UriBuilder()
            {
                Host = ConfigINI.Items["FABPLAN_HOST_NAME"].ToString(),
                Path = ConfigINI.Items["FABPLAN_LOGIN_URL"].ToString(),
                Port = -1,
                Scheme = "http"
            };

            dynamic credentials = new
            {
                username = ConfigINI.Items["FABPLAN_USER_NAME"],
                password = ConfigINI.Items["FABPLAN_PASSWORD"]
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

        /// <summary>
        /// Logs out from Fabplan.
        /// </summary>
        private async Task LogOutFromFabplan()
        {
            UriBuilder builder = new UriBuilder()
            {
                Host = ConfigINI.Items["FABPLAN_HOST_NAME"].ToString(),
                Path = ConfigINI.Items["FABPLAN_LOGOUT_URL"].ToString(),
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
