using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Text;
using System.Globalization;
using ExampleBase;

using NetOffice;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace WordExamplesCS4
{
    /// <summary>
    /// Example 4 - Using a datasource
    /// </summary>
    internal class Example04 : IExample
    {
        #region IExample

        public void RunExample()
        {
            // create simple a csv-file as datasource
            string fileName = string.Format("{0}\\DataSource.csv", HostApplication.RootDirectory);
             
            // if file exists then delete
            if (File.Exists(fileName))
                File.Delete(fileName);

            File.AppendAllText(fileName, string.Format("{0},{1}{2}", "ProjectName", "ProjectLink", Environment.NewLine));
            File.AppendAllText(fileName, string.Format("{0},{1}{2}", "NetOffice", "http://netoffice.codeplex.com", Environment.NewLine));

            // start word and turn off msg boxes
            Word.Application wordApplication = new Word.Application();
            wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            // create a utils instance, not need for but helpful to keep the lines of code low
            Word.Tools.CommonUtils utils = new Word.Tools.CommonUtils(wordApplication);

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

            object adress = newDocument.MailMerge.DataSource.DataFields[2].Value;
            object screenTip = "click me if you want!";
            object displayText = "here";
            newDocument.Hyperlinks.Add(wordApplication.Selection.Range, adress, Missing.Value, screenTip, displayText, Missing.Value);

            // show the contents of the fields
            int wdToggle = 9999998;
            newDocument.MailMerge.ViewMailMergeFieldCodes = wdToggle;

            //do not show the fieldcodes
            wordApplication.ActiveWindow.View.ShowFieldCodes = false;

            // save the document
            string documentFile = utils.File.Combine(HostApplication.RootDirectory, "Example04", Word.Tools.DocumentFormat.Normal);
            newDocument.SaveAs(documentFile);

            // close word and dispose reference
            wordApplication.Quit();
            wordApplication.Dispose();

            // show dialog for the user(you!)
            HostApplication.ShowFinishDialog(null, documentFile);
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public string Caption
        {
            get { return HostApplication.LCID == 1033 ? "Example04" : "Beispiel04"; }
        }

        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Using data source" : "Verwendung von DataSource"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion

        #region Properties

        internal IHost HostApplication { get; private set; }

        #endregion
    }
}
