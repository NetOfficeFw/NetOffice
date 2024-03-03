using System;
using System.Collections.Generic;
using Excel = NetOffice.ExcelApi;
using System.Windows.Forms;
using NetOffice.ExcelApi.Tools;

namespace Sample.ExcelAddin
{
    /// <summary>
    /// Custom pane for MS-Excel. The control implements the ITaskPane interface from NetOffice.ExcelApi.Tools(no need for but helpful)
    /// </summary>
    public partial class TranslationPane : UserControl, ITaskPane
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public TranslationPane()
        {
            try
            {
                InitializeComponent();
                InitializeTranslationClient();              
            }
            catch (Exception exception)
            {
                ShowError(string.Format("An errror occured. Details: {0}", exception.Message));
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// IPC Client Proxy to the local Translation Server
        /// </summary>
        internal TranslationClient Client { get; set; }
         
        /// <summary>
        /// HostInstance
        /// </summary>
        private Excel.Application Application { get; set; }

        #endregion

        #region UI Trigger

        private void buttonTranslate_Click(object sender, EventArgs e)
        {
            try
            {
                ClearError();
                string translatedText = Client.DataService.Translate(
                                                     comboBoxSourceLanguage.SelectedItem as string,
                                                     comboBoxTargetLanguage.SelectedItem as string,
                                                     textBoxRequested.Text);
                textBoxTranslation.Text = translatedText;
            }
            catch (Exception exception)
            {
                ShowError(string.Format("An errror occured. Details: {0}",exception.Message));  
            }         
        }

        private void textBoxRequested_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                buttonTranslate_Click(buttonTranslate, new EventArgs());
                e.Handled = true;
            }
        }

        private void linkLabelNetOfficePage_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("http://netoffice.codeplex.com");
            }
            catch
            {
                ; // no problem ;)
            }
        }

        private void buttonDoLocalConnect_Click(object sender, EventArgs e)
        {
            try
            {
                ClearError();
                InitializeTranslationClient();
            }
            catch (Exception exception)
            {
                ShowError(string.Format("An errror occured. Details: {0}", exception.Message));
            }
        }

        #endregion

        #region Excel Trigger

        /// <summary>
        /// This method was called from the host instance after selection changed.
        /// </summary>
        /// <param name="selectedRange">active selection</param>
        private void Application_SheetSelectionChangeEvent(NetOffice.ICOMObject sh, Excel.Range selectedRange)
        {
            try
            {
                // we check for auto translation and skip selections with more than 256 cells for this simple example
                if (checkBoxAutoTranslate.Checked && selectedRange.Count <= 256)
                {
                    ClearError();
                    textBoxTranslation.Text = string.Empty;
                    foreach (var item in selectedRange)
                    {
                        string requestedText = item.Text as string;
                        if (!String.IsNullOrWhiteSpace(requestedText))
                        {
                            string translatedText = Client.DataService.Translate(
                                                         comboBoxSourceLanguage.SelectedItem as string,
                                                         comboBoxTargetLanguage.SelectedItem as string,
                                                         requestedText);
                            textBoxTranslation.Text += translatedText + " ";
                        }
                    }
                }

                sh.Dispose();
                selectedRange.Dispose();
            }
            catch (Exception exception)
            {
                ShowError(string.Format("An errror occured. Details: {0}", exception.Message));
            }
        }

        #endregion

        #region ITaskPane Member

        public void OnConnection(Excel.Application application, NetOffice.OfficeApi._CustomTaskPane parentPane, object[] customArguments)
        {
            try
            {
                Application = application;
                Application.SheetSelectionChangeEvent += new Excel.Application_SheetSelectionChangeEventHandler(Application_SheetSelectionChangeEvent);
            }
            catch (Exception exception)
            {
                ShowError(string.Format("An errror occured. Details: {0}", exception.Message));
            }
        }

        public void OnDisconnection()
        {

        }

        public void OnDockPositionChanged(NetOffice.OfficeApi.Enums.MsoCTPDockPosition position)
        {

        }

        public void OnVisibleStateChanged(bool visible)
        {

        }

        #endregion

        #region Methods

        /// <summary>
        /// Initialize the local translation client
        /// </summary>
        private void InitializeTranslationClient()
        {
            // create the IPC proxy to the translation service and initialize the panel
            if (null != Client)
                Client.Dispose();
            Client = new TranslationClient();
            comboBoxSourceLanguage.DataSource = Client.DataService.AvailableTranslations;
            comboBoxTargetLanguage.DataSource = Client.DataService.AvailableTranslations;
            comboBoxSourceLanguage.SelectedItem = "English";
            comboBoxTargetLanguage.SelectedItem = "German";
        }

        /// <summary>
        /// Clear Error Panel
        /// </summary>
        private void ClearError()
        {
            labelErrorMessage.Text = string.Empty;
            panelError.Visible = false;
        }

        /// <summary>
        /// Display Error Message
        /// </summary>
        /// <param name="errorMessage">error message to display</param>
        private void ShowError(string errorMessage)
        {
            pictureBoxError.Image = pictureBoxInitial.Image;
            pictureBoxError.Visible = true;
            labelErrorMessage.ForeColor = System.Drawing.Color.Red;
            labelErrorMessage.Text = errorMessage;
            panelError.Visible = true;
        }

        #endregion
    }
}
