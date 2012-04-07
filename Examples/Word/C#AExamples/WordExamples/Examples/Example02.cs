using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.Text;
using ExampleBase;

using LateBindingApi.Core;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace WordExamples
{
    class Example02 : IExample
    {
        IHost _hostApplication;

        #region IExample Member

        public void RunExample()
        {
            // Initialize NetOffice
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

            // we save the document as .doc for compatibility with all word versions
            string documentFile = string.Format("{0}\\Example02{1}", _hostApplication.RootDirectory, ".doc");
            newDocument.SaveAs(documentFile, WdSaveFormat.wdFormatDocumentDefault);

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
            get { return _hostApplication.LCID == 1033 ? "Example02" : "Beispiel02"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Insert a table to document" : "Eine Tabelle erstellen"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion
    }
}
