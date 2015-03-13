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
        /// Event Repeater for state updates
        /// </summary>
        private DataEventRepeator Repeater { get; set; }

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
        /// Initialize the local translation in application
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
        /// This method was called from the IPC Service instance for state updates
        /// This is a very stupid solution because the service is calling the view directly and wait until its finished
        /// Never do this in a real-life scenario because the view can block/slow down the service
        /// </summary>
        /// <param name="result">operation state</param>
        private void ServiceOnTranslation(TranslateOperationResult result)
        {
            if (this.InvokeRequired)
            {
                Action<TranslateOperationResult> invoker = ServiceOnTranslation;
                invoker.Invoke(result);
            }
            else
            {
                switch (result.State)
                {
                    case TranslateOperationState.Sucseed:
                        DisplayTranslationMessage(result.Requested, result.Result, result.Cached);
                        break;
                    case TranslateOperationState.Error:
                        DisplayFailedMessage(result.Exception);
                        break;
                    default:
                        throw new IndexOutOfRangeException("state");
                }
            }
        }

        /// <summary>
        /// Display translation sucseed message
        /// </summary>
        /// <param name="request">requested text</param>
        /// <param name="response">response from translations</param>
        /// <param name="fromCache">reponse re-used from local cache</param>
        private void DisplayTranslationMessage(string request, string response, bool fromCache)
        {
            string displayRequest = (request.Length > 24 ? request.Substring(0, 21) + "..." : request).Replace(Environment.NewLine, " ");
            string displayResponse = (response.Length > 24 ? response.Substring(0, 21) + "..." : response).Replace(Environment.NewLine, " ");
            string message = String.Format("Translation \"{0}\"=>\"{1}\"; FromCache={2}", displayRequest, displayResponse, fromCache);

            ListViewItem item = new ListViewItem();
            item.ImageIndex = fromCache ? 1 : 2;
            item.Text = DateTime.Now.ToString();
            item.SubItems.Add(message);
            listView1.Items.Insert(0, item);
            ValidateViewItemsCount();
        }

        /// <summary>
        /// Display failed translation
        /// </summary>
        /// <param name="exception">origin exception</param>
        private void DisplayFailedMessage(Exception exception)
        {
            string message = String.Format("Translation failed because {0}", exception);
            ListViewItem item = new ListViewItem();
            item.ImageIndex = 3;
            item.Text = DateTime.Now.ToString();
            item.SubItems.Add(message);
            listView1.Items.Insert(0, item);
            ValidateViewItemsCount();
        }

        /// <summary>
        /// Say hello to the user(or anyone else)
        /// </summary>
        private void DisplayStartMessage()
        {
            listView1.Items.Clear();
            ListViewItem item = listView1.Items.Add(DateTime.Now.ToString());
            item.SubItems.Add("Server application started. Use Translation from MS-Excel Addin or Translation Tab.");
            item.ImageIndex = 0;
            ValidateViewItemsCount();
        }

        /// <summary>
        /// Make sure to display the last 256 messages only
        /// </summary>
        private void ValidateViewItemsCount()
        {
            int overheadCount = listView1.Items.Count - 256;
            if (overheadCount > 0)
            {
                for (int i = 0; i < overheadCount; i++)
                    listView1.Items.RemoveAt(listView1.Items.Count - 1);
            }
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

        private void buttonClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
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

        private void textBoxRequested_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if(e.KeyCode == Keys.Return && e.Control)
                    buttonTranslate_Click(buttonTranslate, EventArgs.Empty);
            }
            catch (Exception exception)
            {
                MessageBox.Show(String.Format("An errror occured. Details: {0}", exception.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
