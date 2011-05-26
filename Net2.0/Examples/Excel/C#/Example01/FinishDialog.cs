using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using System.Diagnostics;

namespace Example01
{
    public partial class FinishDialog : Form
    {

        string _message;
        string _DocumentPath;

        public FinishDialog(string message, string documentPath)
        {
            InitializeComponent();

            _message = message;
            _DocumentPath = documentPath;

            labelMessage.Text = _message;
            labelDocumentPath.Text = _DocumentPath;
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonOpenDocument_Click(object sender, EventArgs e)
        {
            Process.Start(_DocumentPath);
            this.Close();
        }
    }
}
