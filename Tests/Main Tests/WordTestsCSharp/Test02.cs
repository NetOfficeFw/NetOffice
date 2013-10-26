using System;
using System.Collections.Generic;
using System.Text;
using Tests.Core;
using NetOffice;
using Office = NetOffice.OfficeApi;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace WordTestsCSharp
{
    public class Test02 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test02"; }
        }

        public string Description
        {
            get { return "Using a DataTable"; }
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

                Word.Document newDocument = application.Documents.Add();
                Word.Table table = newDocument.Tables.Add(application.Selection.Range, 3, 2);

                // insert some text into the cells
                table.Cell(1, 1).Select();
                application.Selection.TypeText("This");

                table.Cell(1, 2).Select();
                application.Selection.TypeText("table");

                table.Cell(2, 1).Select();
                application.Selection.TypeText("was");

                table.Cell(2, 2).Select();
                application.Selection.TypeText("created");

                table.Cell(3, 1).Select();
                application.Selection.TypeText("by");

                table.Cell(3, 2).Select();
                application.Selection.TypeText("NetOffice");

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
