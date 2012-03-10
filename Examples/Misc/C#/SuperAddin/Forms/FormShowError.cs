using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace SuperAddinCSharp
{
    partial class FormShowError : Form
    {
        #region Fields

        private string _errorHeader;
        private string _errorFooter;
        private Exception _exception;
        private bool _isExtended;

        #endregion

        #region Construction

        public FormShowError(Exception exceptionToShow)
        {
            InitializeComponent();

            Initialize("An error is occured.", "", exceptionToShow);
            labelErrorHeader.Visible = true;

        }

        public FormShowError(Exception exceptionToShow, string errorFooter)
        {
            InitializeComponent();

            Initialize("An error is occured.", errorFooter, exceptionToShow);
            labelErrorHeader.Visible = true;
            labelErrorFooter.Visible = true;
        }

        public FormShowError(string errorHeader, string errorFooter, Exception exceptionToShow)
        {
            InitializeComponent();

            Initialize(errorHeader, errorFooter, exceptionToShow);
            labelErrorHeader.Visible = true;
            labelErrorFooter.Visible = true;
        }

        private void Initialize(string errorHeader, string errorFooter, Exception exceptionToShow)
        {
            this.Width = 440;
            this.Height = 160;
            _isExtended = false;

            _errorHeader = errorHeader;
            _errorFooter = errorFooter;

            labelErrorHeader.Text = errorHeader;
            labelErrorFooter.Text = errorFooter;

            _exception = exceptionToShow;
            int i = 1;
            while (_exception != null)
            {
                ListViewItem lviException = listViewExceptions.Items.Insert(0, i.ToString());
                lviException.SubItems.Add(_exception.Source);
                lviException.SubItems.Add(exceptionToShow.GetType().ToString());
                lviException.SubItems.Add(_exception.Message);
                _exception = _exception.InnerException;
                i++;
            }
            _exception = exceptionToShow;

            listViewExceptions_Resize(this, new EventArgs());

        }

        #endregion

        #region GuiTrigger

        private void listViewExceptions_Resize(object sender, EventArgs e)
        {
            listViewExceptions.Columns[0].Width = (listViewExceptions.Width / 100) * 10;
            listViewExceptions.Columns[1].Width = (listViewExceptions.Width / 100) * 20;
            listViewExceptions.Columns[2].Width = (listViewExceptions.Width / 100) * 20;
            listViewExceptions.Columns[3].Width = (listViewExceptions.Width / 100) * 60;
        }

        private void buttonDetails_Click(object sender, EventArgs e)
        {

            if (true == _isExtended)
            {
                this.Width = 440;
                this.Height = 160;
                _isExtended = false;
            }
            else
            {
                this.Width = 440;
                this.Height = 300;
                _isExtended = true;
            }
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #endregion

        /// <summary>
        /// writes message to logfile in dll folder
        /// </summary>
        /// <param name="header"></param>
        /// <param name="throwedException"></param>
        public static void LogError(string header, Exception throwedException)
        {
            // dll path
            string codeBase = Assembly.GetCallingAssembly().CodeBase;
            if (true == codeBase.StartsWith("file:///", StringComparison.InvariantCultureIgnoreCase))
                codeBase = codeBase.Substring(8);
            codeBase = codeBase.Replace(@"/", @"\");
            codeBase = codeBase.Substring(0,codeBase.LastIndexOf(@"\"));

            // write message
            string logPath = System.IO.Path.Combine(codeBase, "SuperAddinErrors.log");
            string message ="";
            if(null !=throwedException)
                message = throwedException.Message + "\r\n";
           
            System.IO.File.AppendAllText(logPath, header + "\r\n" + message);
        }
    }
}
