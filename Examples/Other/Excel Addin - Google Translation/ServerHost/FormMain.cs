using System;
using System.Collections.Generic;
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
                InitializeTranslationTab();
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
        /// Initialize the local Localization
        /// </summary>
        private void InitializeTranslationTab()
        {
            // we need a copy from the available languages to fool the currency manager
            comboBoxSourceLanguage.DataSource = CopyStringArray(Service.AvailableTranslations);
            comboBoxTargetLanguage.DataSource = CopyStringArray(Service.AvailableTranslations);
            comboBoxSourceLanguage.SelectedItem = "English";
            comboBoxTargetLanguage.SelectedItem = "German";
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
            item.SubItems.Add("Server application started. Use Translation from MS-Excel Addin or Translation Tab.");
            item.ImageIndex = 0;
        }

        /// <summary>
        /// Creates an array copy
        /// </summary>
        /// <param name="array">array to copy them</param>
        /// <returns>deep array clone</returns>
        private static string[] CopyStringArray(string[] array)
        {
            List<string> list = new List<string>();
            foreach (string item in array)
                list.Add(item);
            return list.ToArray();
        }

        private void buttonTranslate_Click(object sender, EventArgs e)
        {
            try
            {
                string translatedText = Service.Translate(
                                                     comboBoxSourceLanguage.SelectedItem as string,
                                                     comboBoxTargetLanguage.SelectedItem as string,
                                                     textBoxRequested.Text);
                textBoxTranslation.Text = translatedText;
            }
            catch (Exception exception)
            {
                MessageBox.Show(String.Format("An errror occured. Details: {0}", exception.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }         
        }
    }
}
