using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ExampleBase
{
    /// <summary>
    /// Example finish dialog
    /// </summary>
    public partial class FinishForm : Form
    {
        #region Fields

        private string _documentPath;
        
        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">finish message</param>
        /// <param name="documentPath">path to created document</param>
        public FinishForm(string message, string documentPath)
        {
            InitializeComponent();
            if (null == message)
                message = "Document saved";

            labelMessage.Text = message;
            labelDocumentPath.Text = documentPath;
            _documentPath = documentPath;
        }

        #endregion

        #region Methods

        private void TryDeleteDocument()
        {
            try
            {
                if (File.Exists(_documentPath))
                    File.Delete(_documentPath);
            }
            catch
            {
                // still open
                Console.WriteLine("Unable to delete {0}.", _documentPath);
            }
        }

        #endregion

        #region Trigger

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
                if (checkBoxDeleteDocument.Checked)
                    TryDeleteDocument();
                this.Close();
            }
            catch
            {
                MessageBox.Show(this, "An error occured.", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion
    }
}
