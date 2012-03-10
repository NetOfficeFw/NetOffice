using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using NetOffice.VBIDEApi.Enums;

namespace Example02
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize Api COMObject Support
            LateBindingApi.Core.Factory.Initialize();

            // start word and turn off msg boxes
            Word.Application wordApplication = new Word.Application();
            wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            // add a new document
            Word.Document newDocument = wordApplication.Documents.Add();

            // add a table
            Word.Table table = newDocument.Tables.Add(wordApplication.Selection.Range, 3, 2);

            // insert some text into the cells
            table.Cell(1, 1).Select();
            wordApplication.Selection.TypeText("This");

            table.Cell(1, 2).Select();
            wordApplication.Selection.TypeText("table");

            table.Cell(2, 1).Select();
            wordApplication.Selection.TypeText("was");

            table.Cell(2, 2).Select();
            wordApplication.Selection.TypeText("created");

            table.Cell(3, 1).Select();
            wordApplication.Selection.TypeText("by");

            table.Cell(3, 2).Select();
            wordApplication.Selection.TypeText("NetOffice");

            // save the document
            string fileExtension = GetDefaultExtension(wordApplication);
            object documentFile = string.Format("{0}\\Example02{1}", Application.StartupPath, fileExtension);
            newDocument.SaveAs(documentFile);

            // close word and dispose reference
            wordApplication.Quit();
            wordApplication.Dispose();

            FinishDialog fDialog = new FinishDialog("Document saved.", documentFile.ToString());
            fDialog.ShowDialog(this);
        }

        #region Helper

        /// <summary>
        /// returns the valid file extension for the instance. for example ".doc" or ".docx"
        /// </summary>
        /// <param name="application">the instance</param>
        /// <returns>the extension</returns>
        private static string GetDefaultExtension(Word.Application application)
        {
            double version = Convert.ToDouble(application.Version);
            if (version >= 120.00)
                return ".docx";
            else
                return ".doc";
        }

        #endregion
    }
}
