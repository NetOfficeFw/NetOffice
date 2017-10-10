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
    /// Example 1 - Create a document
    /// </summary>
    internal class Example01 : IExample
    {
        public void RunExample()
        {
            // start word and turn off msg boxes
            Word.Application wordApplication = new Word.Application();
            wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            // create a utils instance, no need for but helpful to keep the lines of code low
            CommonUtils utils = new CommonUtils(wordApplication);

            // add a new document
            Word.Document newDocument = wordApplication.Documents.Add();

            // insert some text
            wordApplication.Selection.TypeText("This text is written by automation");

            wordApplication.Selection.HomeKey(WdUnits.wdLine, WdMovementType.wdExtend);
            wordApplication.Selection.Font.Color = WdColor.wdColorSeaGreen;
            wordApplication.Selection.Font.Bold = 1;
            wordApplication.Selection.Font.Size = 18;

            // save the document
            string documentFile = utils.File.Combine(HostApplication.RootDirectory, "Example01", DocumentFormat.Normal);
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
            get { return "Example01"; }
        }

        public string Description
        {
            get { return "Create a document write text and save"; }
        }

        public UserControl Panel 
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}
