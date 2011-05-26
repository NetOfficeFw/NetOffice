using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Diagnostics;
using System.Windows.Forms;

namespace Example02
{
    public partial class FinishDialog : Form
    {
        
        string _message;
        string _documentPath;

        public FinishDialog(string message, string documentPath)
        {
            InitializeComponent();

            _message = message;
            _documentPath = documentPath;

            labelMessage.Text      = _message;
            labelDocumentPath.Text = _documentPath; 
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonOpenDocument_Click(object sender, EventArgs e)
        {
            Process.Start(_documentPath);         
            this.Close();
        }
    }
}
