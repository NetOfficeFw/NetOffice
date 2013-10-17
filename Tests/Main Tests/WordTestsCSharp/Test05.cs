using System;
using System.Collections.Generic;
using System.Text;
using Tests.Core;
using NetOffice;
using Office = NetOffice.OfficeApi;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using NetOffice.VBIDEApi.Enums;

namespace WordTestsCSharp
{
    public class Test05 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test05"; }
        }

        public string Description
        {
            get { return "Using a VBE Macros."; }
        }

        public string OfficeProduct
        {
            get { return "Word"; }
        }

        public string Language
        {
            get { return "C#"; }
        }

        public TestResult DoTest()
        {
            Word.Application application = null;
            DateTime startTime = DateTime.Now;
            try
            {
                application = new Word.Application();
                application.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                application.Visible = true;

                Word.Document newDocument = application.Documents.Add();

                // add new module and insert macro
                // the option "Trust access to Visual Basic Project" must be set
                NetOffice.VBIDEApi.CodeModule module = newDocument.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule).CodeModule;

                // set the modulename
                module.Name = "NetOfficeTestModule";

                //add the macro
                string codeLines = string.Format("Public Sub NetOfficeTestMacro()\r\n   {0}\r\nEnd Sub",
                                                 "Selection.TypeText (\"This text is written by a automatic created macro with NetOffice...\")");
                module.InsertLines(1, codeLines);

                //start the macro NetOfficeTestModule
                application.Run("NetOfficeTestModule!NetOfficeTestMacro");
               
                newDocument.Close(false);

                return new TestResult(true, DateTime.Now.Subtract(startTime), "", null, "");
            }
            catch (Exception exception)
            {
                return new TestResult(false, DateTime.Now.Subtract(startTime), exception.Message, exception, "");
            }
            finally
            {
                if (null != application)
                {
                    application.Quit(WdSaveOptions.wdDoNotSaveChanges);
                    application.Dispose();
                }
            }
        }

        #endregion
    }
}
