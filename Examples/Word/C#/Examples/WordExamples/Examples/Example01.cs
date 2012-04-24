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
    class Example01 : IExample
    {
        IHost _hostApplication;
        
        #region IExample Member

        public void RunExample()
        {
            // start word and turn off msg boxes
            Word.Application wordApplication = new Word.Application();
            wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            // add a new document
            Word.Document newDocument = wordApplication.Documents.Add();

            // insert some text
            wordApplication.Selection.TypeText("This text is written by NetOffice");

            wordApplication.Selection.HomeKey(WdUnits.wdLine, WdMovementType.wdExtend);
            wordApplication.Selection.Font.Color = WdColor.wdColorSeaGreen;
            wordApplication.Selection.Font.Bold = 1;
            wordApplication.Selection.Font.Size = 18;

            // we save the document as .doc for compatibility with all word versions
            string documentFile = string.Format("{0}\\Example01{1}", _hostApplication.RootDirectory, ".doc");
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
            get { return _hostApplication.LCID == 1033 ? "Example01" : "Beispiel01"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Create a document write text and save" : "Dokument erstellen, Text schreiben und speichern"; }
        }

        public UserControl Panel 
        {
            get { return null; }
        }

        #endregion
    }
}
