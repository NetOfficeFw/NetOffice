using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Text;
using System.Globalization;
using ExampleBase;

using LateBindingApi.Core;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace WordExamplesCS4
{
    class Example04 : IExample
    {
        IHost _hostApplication;

        #region IExample Member

        public void RunExample()
        {
            // Initialize NetOffice
            LateBindingApi.Core.Factory.Initialize();

            // create simple a csv-file as datasource
            string fileName = string.Format("{0}\\DataSource.csv", _hostApplication.RootDirectory);
             
            // if file exists then delete
            if (File.Exists(fileName))
                File.Delete(fileName);

            File.AppendAllText(fileName, string.Format("{0},{1}{2}", "ProjectName", "ProjectLink", Environment.NewLine));
            File.AppendAllText(fileName, string.Format("{0},{1}{2}", "NetOffice", "http://netoffice.codeplex.com", Environment.NewLine));

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

            object adress = newDocument.MailMerge.DataSource.DataFields[2].Value;
            object screenTip = "click me if you want!";
            object displayText = "here";
            newDocument.Hyperlinks.Add(wordApplication.Selection.Range, adress, Missing.Value, screenTip, displayText, Missing.Value);

            // show the contents of the fields
            int wdToggle = 9999998;
            newDocument.MailMerge.ViewMailMergeFieldCodes = wdToggle;

            //do not show the fieldcodes
            wordApplication.ActiveWindow.View.ShowFieldCodes = false;

            // we save the document as .doc for compatibility with all word versions
            string documentFile = string.Format("{0}\\Example04{1}", _hostApplication.RootDirectory, ".doc");
            double wordVersion = Convert.ToDouble(wordApplication.Version, CultureInfo.InvariantCulture);
            if (wordVersion >= 12.0)
                newDocument.SaveAs(documentFile, WdSaveFormat.wdFormatDocumentDefault);
            else
                newDocument.SaveAs(documentFile);

            // close word and dispose reference
            wordApplication.Quit();
            wordApplication.Dispose();

            // show dialog for the user(you!)
            _hostApplication.ShowFinishDialog(null, documentFile);
        }

        public void Connect(IHost hostApplication)
        {
            _hostApplication = hostApplication;
        }

        public string Caption
        {
            get { return _hostApplication.LCID == 1033 ? "Example04" : "Beispiel04"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Using data source" : "Verwendung von DataSource"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion
    }
}
