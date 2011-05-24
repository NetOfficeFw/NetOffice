using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums; 

namespace Example04
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

            // Initialize Api COMObject Support
            LateBindingApi.Core.Factory.Initialize();

            // create simple a csv-file as datasource
            string fileName = string.Format("{0}\\DataSource.csv", Application.StartupPath);

            // if file exists then delete
            if (File.Exists(fileName)) 
                File.Delete(fileName);

            File.AppendAllText(fileName, string.Format("{0},{1}{2}", "ProjectName", "ProjectLink", Environment.NewLine));
            File.AppendAllText(fileName, string.Format("{0},{1}{2}", "NetOffice", "http://netoffice.codeplex.com/", Environment.NewLine));

            // start word and turn off msg boxes
            Word.Application wordApplication = new Word.Application();
            wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            // add a new document
            Word.Document newDocument = wordApplication.Documents.Add();

            // define the document as mailmerge
            newDocument.MailMerge.MainDocumentType = WdMailMergeMainDocType.wdFormLetters;

            // open the datasource
            newDocument.MailMerge.OpenDataSource(fileName);

            // insert some text and the mailmergefields defined in the datasource
            wordApplication.Selection.TypeText("This test is brought to you by ");
            newDocument.MailMerge.Fields.Add(wordApplication.Selection.Range, "ProjectName");

            wordApplication.Selection.TypeText(" for more information and examples visit ");
            newDocument.MailMerge.Fields.Add(wordApplication.Selection.Range, "ProjectLink ");

            wordApplication.Selection.TypeText(" or click ");

            object adress = newDocument.MailMerge.DataSource.DataFields[1].Value;
            object screenTip = "come on dude click me, i know you want it...";
            object displayText = "here";
            newDocument.Hyperlinks.Add(wordApplication.Selection.Range, adress, Missing.Value, screenTip, displayText, Missing.Value);

            // show the contents of the fields
            int wdToggle = 9999998;
            newDocument.MailMerge.ViewMailMergeFieldCodes = wdToggle;

            //do not show the fieldcodes
            wordApplication.ActiveWindow.View.ShowFieldCodes = false;

            //save the document
            string fileExtension = GetDefaultExtension(wordApplication);
            object documentFile = string.Format("{0}\\Example04{1}", Application.StartupPath, fileExtension);
            newDocument.SaveAs(documentFile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            //close word and dispose reference
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
