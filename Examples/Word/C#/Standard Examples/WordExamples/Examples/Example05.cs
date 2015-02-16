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
using VB = NetOffice.VBIDEApi;
using NetOffice.VBIDEApi.Enums;

namespace WordExamplesCS4
{
    /// <summary>
    /// Example 5 - Create macros
    /// </summary>
    internal class Example05 : IExample
    {
        #region IExample Member

        public void RunExample()
        {
            // start word and turn off msg boxes
            Word.Application wordApplication = new Word.Application();
            wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            wordApplication.Visible = true;

            // add a new document
            Word.Document newDocument = wordApplication.Documents.Add();

            // add new module and insert macro
            // the option "Trust access to Visual Basic Project" must be set
            VB.CodeModule module = newDocument.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule).CodeModule;

            // set the modulename
            module.Name = "NetOfficeTestModule";

            //add the macro
            string codeLines = string.Format("Public Sub NetOfficeTestMacro()\r\n   {0}\r\nEnd Sub",
                                             "Selection.TypeText (\"This text is written by a automatic created macro with NetOffice...\")");
            module.InsertLines(1, codeLines);

            //start the macro NetOfficeTestModule
            wordApplication.Run("NetOfficeTestModule!NetOfficeTestMacro");

            string fileExtension = GetFileExtension(wordApplication);
            string documentFile = string.Format("{0}\\Example05{1}", HostApplication.RootDirectory, fileExtension);
            double wordVersion = Convert.ToDouble(wordApplication.Version, CultureInfo.InvariantCulture);
            if (wordVersion >= 12.0)
                newDocument.SaveAs(documentFile, WdSaveFormat.wdFormatXMLDocumentMacroEnabled);
            else
                newDocument.SaveAs(documentFile);

            //close word and dispose reference
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
            get { return HostApplication.LCID == 1033 ? "Example05" : "Beispiel05"; }
        }

        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Create vba macros. The option Trust access to Visual Basic Project must be set." : "Erstellen von VBA Macros. Die Option Visual Basic Projekten vertrauen muss aktiviert sein."; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion

        #region Properties

        internal IHost HostApplication { get; private set; }

        #endregion

        #region Helper

        string GetFileExtension(Word.Application application)
        {
            double wordVersion = Convert.ToDouble(application.Version, CultureInfo.InvariantCulture);
            if (wordVersion >= 12.0)
                return ".docm";
            else
                return ".docm";
        }

        #endregion
    }
}

