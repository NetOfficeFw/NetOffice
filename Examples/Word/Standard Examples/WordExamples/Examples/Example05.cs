using System;
using System.Windows.Forms;
using ExampleBase;
using NetOffice;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using VB = NetOffice.VBIDEApi;
using NetOffice.VBIDEApi.Enums;
using NetOffice.WordApi.Tools.Contribution;

namespace WordExamplesCS4
{
    /// <summary>
    /// Example 5 - Create macros
    /// </summary>
    internal class Example05 : IExample
    {
        public void RunExample()
        {
            // start word and turn off msg boxes
            Word.Application wordApplication = new Word.Application();
            wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            wordApplication.Visible = true;

            // create a utils instance, not need for but helpful to keep the lines of code low
            CommonUtils utils = new CommonUtils(wordApplication);

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

            // save the document
            string documentFile = utils.File.Combine(HostApplication.RootDirectory, "Example05", DocumentFormat.Macros);
            newDocument.SaveAs(documentFile);
            if(utils.ApplicationIs2007OrHigher)
                newDocument.SaveAs(documentFile, WdSaveFormat.wdFormatXMLDocumentMacroEnabled);
            else
                newDocument.SaveAs(documentFile);

            //close word and dispose reference
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
            get { return "Example05"; }
        }

        public string Description
        {
            get { return "Create vba macros. The option Trust access to Visual Basic Project must be set."; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}

