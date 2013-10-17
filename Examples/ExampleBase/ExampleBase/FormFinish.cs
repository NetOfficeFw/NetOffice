using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ExampleBase
{
    public partial class FormFinish : Form
    {
        string _documentPath;

        public FormFinish(string message, string documentPath)
        {
            InitializeComponent();

            if (null == message)
                message = "Document saved.";

            labelMessage.Text = message;
            labelDocumentPath.Text = documentPath;
            _documentPath = documentPath;
        }

        private void buttonOpenDocument_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(_documentPath);
                this.Close();
            }
            catch
            {
                MessageBox.Show(this, "An error occured.", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            try
            {
                this.Close();
            }
            catch
            {
                MessageBox.Show(this, "An error occured.", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
