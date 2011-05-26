using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using System.Diagnostics;

namespace Example07
{
    partial class FinishDialog : Form
    {
        string _message;
        string _workbookPath;

        public FinishDialog(string message, string workbookPath)
        {
            InitializeComponent();

            _message = message;
            _workbookPath = workbookPath;

            labelMessage.Text = _message;
            labelWorkbookPath.Text = _workbookPath;
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonOpenWorkbook_Click(object sender, EventArgs e)
        {
            Process.Start(_workbookPath);
            this.Close();
        }

    }
}
