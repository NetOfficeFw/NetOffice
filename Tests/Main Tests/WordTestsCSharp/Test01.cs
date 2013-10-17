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
    public class Test01 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test01"; }
        }

        public string Description
        {
            get { return "Insert text."; }
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
                application.Selection.TypeText("This text is written by NetOffice");

                application.Selection.HomeKey(WdUnits.wdLine, WdMovementType.wdExtend);
                application.Selection.Font.Color = WdColor.wdColorSeaGreen;
                application.Selection.Font.Bold = 1;
                application.Selection.Font.Size = 18;

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
