using System;
using System.Windows.Forms;
using System.Runtime.Remoting;
using System.Runtime.Remoting.Channels;
using System.Runtime.Remoting.Channels.Ipc;
using Sample.Server;

namespace Sample.ServerHost
{
    /// <summary>
    /// Main Form for the application
    /// </summary>
    public partial class FormMain : Form
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public FormMain()
        {
            InitializeComponent();
            if (!IsRunning())
            {
                RegisterServer();
                DisplayStartMessage();
            }
            else
            {
                MessageBox.Show("Server is already running.", "Sample.ServerHost", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Application.Exit();
            }
        }

        /// <summary>
        /// The Translation Service Instance
        /// </summary>
        private WebTranslationService Service { get; set; }

        /// <summary>
        /// Event Repeater for status updates
        /// </summary>
        private DataEventRepeator Repeater { get; set; }

        /// <summary>
        /// Temporaily result objekt for UI Invokes
        /// </summary>
        TranslateOperationResult TranslationResult { get; set; }

        /// <summary>
        /// The Close button click trigger
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        /// <summary>
        /// Returns the server is already running
        /// </summary>
        /// <returns></returns>
        private bool IsRunning()
        {
            return null != ChannelServices.GetChannel("NetOffice.SampleChannel");
        }
        
        /// <summary>
        /// Register the Service Property as IPC Server
        /// </summary>
        private void RegisterServer()
        {
            Repeater = new DataEventRepeator();
            IpcServerChannel channel = new IpcServerChannel("NetOffice.SampleChannel");
            ChannelServices.RegisterChannel(channel, true);
            RemotingConfiguration.RegisterWellKnownServiceType(
                                        typeof(WebTranslationService),
                                        "NetOffice.WebTranslationService.DataService",
                                        WellKnownObjectMode.Singleton);

            Service = new WebTranslationService();
            Repeater.Translation += new TranslationEventHandler(ServiceOnTranslation);
            Service.AddEventRepeater(Repeater);
        }
       
        /// <summary>
        /// Helper method to perform a status update in UI thread
        /// </summary>
        private void ServiceOnTranslationInUIThread()
        {
            if (null != TranslationResult)
            { 
                ServiceOnTranslation(TranslationResult);
                TranslationResult = null;
            }
        }

        /// <summary>
        /// This Trigger was called from the IPC Service Instance for status update
        /// </summary>
        /// <param name="result"></param>
        private void ServiceOnTranslation(TranslateOperationResult result)
        {
            if (this.InvokeRequired)
            {
                TranslationResult = result;
                MethodInvoker uiThreadHandler = new MethodInvoker(ServiceOnTranslationInUIThread);
                this.Invoke(uiThreadHandler);
            }
            else
            {
                switch (result.State)
                {
                    case TranslateOperationState.Sucseed:
                        
                        int imageIndex = 1;
                        if(result.Requested.Trim().Equals(result.Result, StringComparison.InvariantCultureIgnoreCase))
                            imageIndex = 2;
                        DisplayMessage(string.Format("Translation sucseed \"{0}\"=>\"{1}\"; FromLocalCache={2}", result.Requested, result.Result, result.Cached), imageIndex);
                        break;
                    case TranslateOperationState.Error:
                        DisplayMessage(string.Format("Translation failed because {0}" , result.Exception.ToString()), 3);
                        break;
                    default:
                        throw new IndexOutOfRangeException("state");
                }
            }
        }

        /// <summary>
        /// Display a new message in listview
        /// </summary>
        /// <param name="message"></param>
        /// <param name="imageIndex"></param>
        private void DisplayMessage(string message, int imageIndex)
        {
            ListViewItem item = new ListViewItem();
            item.ImageIndex = imageIndex;
            item.Text = DateTime.Now.ToString();
            item.SubItems.Add(message);
            listView1.Items.Insert(0, item);
        }

        /// <summary>
        /// Say hello to the User
        /// </summary>
        private void DisplayStartMessage()
        {
            listView1.Items.Clear();
            ListViewItem item = listView1.Items.Add(DateTime.Now.ToString());
            item.SubItems.Add("Server application started. Register the Addin from the VS Solution and start MS-Excel.");
            item.ImageIndex = 0;
        }
    }
}
