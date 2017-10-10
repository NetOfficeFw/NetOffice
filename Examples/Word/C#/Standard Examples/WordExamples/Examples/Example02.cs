using System;
using System.Windows.Forms;
using ExampleBase;
using NetOffice;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using NetOffice.WordApi.Tools.Contribution;

namespace WordExamplesCS4
{
    /// <summary>
    /// Example 2 - Insert table
    /// </summary>
    internal class Example02 : IExample
    {
        public void RunExample()
        {
            // start word and turn off msg boxes
            Word.Application wordApplication = new Word.Application();
            wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            // create a utils instance, not need for but helpful to keep the lines of code low
            CommonUtils utils = new CommonUtils(wordApplication);

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
            string documentFile = utils.File.Combine(HostApplication.RootDirectory, "Example02", DocumentFormat.Normal);
            newDocument.SaveAs(documentFile);

            // close word and dispose reference
            wordApplication.Quit();
            wordApplication.Dispose();

            // show end dialog
            HostApplication.ShowFinishDialog(null, documentFile);
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public string Caption
        {
            get { return "Example02"; }
        }

        public string Description
        {
            get { return "Insert a table to document"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }


        internal IHost HostApplication { get; private set; }
    }
}
