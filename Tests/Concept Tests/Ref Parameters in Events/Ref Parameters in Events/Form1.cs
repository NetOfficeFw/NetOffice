using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using NetOffice;
using Word = NetOffice.WordApi;

namespace Ref_Parameters_in_Events
{
    public partial class Form1 : Form
    {
        Word.Application _application;
        bool _cancelClose = true;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            _application = new Word.Application();
            _application.Visible = true;
            _application.Documents.Add();
            _application.Selection.TypeText("Hello World!");
            _application.DocumentBeforeCloseEvent += new NetOffice.WordApi.Application_DocumentBeforeCloseEventHandler(_application_DocumentBeforeCloseEvent);
            button1.Enabled = false;
            button2.Enabled = true;
            button3.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                // throws an exception if canceled in event
                _application.Documents[1].Close(false);
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);            
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (null != _application)
            {
                _cancelClose = false;
                _application.Documents[1].Close(false);
                _application.Quit();
                _application.Dispose();
                _application = null;
                button1.Enabled = true;
                button2.Enabled = false;
                button3.Enabled = false;
            }

        }

        void _application_DocumentBeforeCloseEvent(NetOffice.WordApi.Document Doc, ref bool Cancel)
        {
            Cancel = _cancelClose;
        }
    }
}
