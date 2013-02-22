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
    public class Test08 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test08"; }
        }

        public string Description
        {
            get { return "Using Paragraphes."; }
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
            Word.Document document = null;
            DateTime startTime = DateTime.Now;
            try
            {
                document = new Word.Document();
                document.Application.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                document.Application.Selection.TypeText("Test with TabIntend C#");
                document.Application.Selection.Start = 0;
                Word.Paragraph p = document.Application.Selection.Range.Paragraphs[1];
               
                p.IndentCharWidth(10);
                p.IndentFirstLineCharWidth(8);
                p.Space1();
                p.Space15();
                p.Space2();
                p.TabHangingIndent(5);
                p.TabIndent(3);

                return new TestResult(true, DateTime.Now.Subtract(startTime), "", null, "");
            }
            catch (Exception exception)
            {
                return new TestResult(false, DateTime.Now.Subtract(startTime), exception.Message, exception, "");
            }
            finally
            {
                if (null != document)
                {
                    document.Application.Quit();
                    document.Dispose();
                }
            }
        }

        #endregion
    }
}
