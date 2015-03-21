using System;
using System.Collections.Generic;
using System.Drawing;
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
    /// Example 1 - Create a document
    /// </summary>
    internal class Example01 : IExample
    {
        #region IExample

        public void RunExample()
        {
            // start word and turn off msg boxes
            Word.Application wordApplication = new Word.Application();
            wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            // create a utils instance, not need for but helpful to keep the lines of code low
            Word.Tools.CommonUtils utils = new Word.Tools.CommonUtils(wordApplication);

            // add a new document
            Word.Document newDocument = wordApplication.Documents.Add();

            // insert some text
            wordApplication.Selection.TypeText("This text is written by NetOffice");

            wordApplication.Selection.HomeKey(WdUnits.wdLine, WdMovementType.wdExtend);
            wordApplication.Selection.Font.Color = WdColor.wdColorSeaGreen;
            wordApplication.Selection.Font.Bold = 1;
            wordApplication.Selection.Font.Size = 18;

            // save the document
            string documentFile = utils.File.Combine(HostApplication.RootDirectory, "Example01", Word.Tools.DocumentFormat.Normal);
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
            get { return HostApplication.LCID == 1033 ? "Example01" : "Beispiel01"; }
        }

        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Create a document write text and save" : "Dokument erstellen, Text schreiben und speichern"; }
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
