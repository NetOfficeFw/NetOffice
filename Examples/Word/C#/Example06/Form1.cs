using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using NetOffice.OfficeApi.Enums;

namespace Example06
{
    public partial class Form1 : Form
    {
        private delegate void UpdateEventTextDelegate(string Message);
        UpdateEventTextDelegate _updateDelegate;

        public Form1()
        {
            InitializeComponent();
            _updateDelegate = new UpdateEventTextDelegate(UpdateTextbox);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize Api COMObject Support
            LateBindingApi.Core.Factory.Initialize();

            // start word and turn off msg boxes
            Word.Application wordApplication = new Word.Application();
            wordApplication.Visible = true;
            /*
            we register some events. note: the event trigger was called from word, means an other Thread
            remove the Quit() call below and check out more events if you want
            */
            wordApplication.NewDocumentEvent += new NetOffice.WordApi.Application_NewDocumentEventHandler(wordApplication_NewDocumentEvent);
            wordApplication.DocumentBeforeCloseEvent += new NetOffice.WordApi.Application_DocumentBeforeCloseEventHandler(wordApplication_DocumentBeforeCloseEvent);

            // add new document and close
            Word.Document document = wordApplication.Documents.Add();
            document.Close();
            
            // close word and dispose reference
            wordApplication.Quit();
            wordApplication.Dispose();
        }

        void wordApplication_DocumentBeforeCloseEvent(NetOffice.WordApi.Document Doc, ref bool Cancel)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event DocumentBeforeClose called." });
            Doc.Dispose();
        }

        void wordApplication_NewDocumentEvent(NetOffice.WordApi.Document Doc)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event NewDocumentEvent called." });
            Doc.Dispose();
        }

        private void UpdateTextbox(string message)
        {
            textBoxEvents.AppendText(message + "\r\n");
        }
    }
}
