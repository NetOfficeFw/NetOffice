using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using ExampleBase;

using NetOffice;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using NetOffice.OfficeApi.Enums;

namespace WordExamplesCS4
{
    internal partial class Example06 : UserControl, IExample
    {
        #region Fields

        private delegate void UpdateEventTextDelegate(string Message);
        private UpdateEventTextDelegate _updateDelegate;

        #endregion

        #region Ctor

        public Example06()
        {
            InitializeComponent();
            _updateDelegate = new UpdateEventTextDelegate(UpdateTextbox);
        }

        #endregion

        #region IExample

        public void RunExample()
        {
            // its an example with an own visual control
            // checkout buttonStartExample_Click
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public string Caption
        {
            get { return HostApplication.LCID == 1033 ? "Example06" : "Beispiel06"; }
        }

        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Using Events" : "Verwenden von Ereignissen"; }
        }

        public UserControl Panel
        {
            get { return this; }
        }

        #endregion

        #region Properties

        internal IHost HostApplication { get; private set; }

        #endregion

        #region Methods
        
        private void UpdateTextbox(string message)
        {
            textBoxEvents.AppendText(message + "\r\n");
        }

        #endregion

        #region Trigger

        private void buttonStartExample_Click(object sender, EventArgs e)
        {
            // start word and turn off msg boxes
            Word.Application wordApplication = new Word.Application();
            wordApplication.Visible = true;            
            
            // we register some events. note: the event trigger was called from word, means an other Thread
            wordApplication.NewDocumentEvent += new NetOffice.WordApi.Application_NewDocumentEventHandler(wordApplication_NewDocumentEvent);
            wordApplication.DocumentBeforeCloseEvent += new NetOffice.WordApi.Application_DocumentBeforeCloseEventHandler(wordApplication_DocumentBeforeCloseEvent);

            // add new document and close
            Word.Document document = wordApplication.Documents.Add();
            document.Close();

            // close word and dispose reference
            wordApplication.Quit();
            wordApplication.Dispose();
        }

        private void wordApplication_DocumentBeforeCloseEvent(NetOffice.WordApi.Document Doc, ref bool Cancel)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event DocumentBeforeClose called." });
            Doc.Dispose();
        }

        private void wordApplication_NewDocumentEvent(NetOffice.WordApi.Document Doc)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event NewDocumentEvent called." });
            Doc.Dispose();
        }

        #endregion
    }
}
