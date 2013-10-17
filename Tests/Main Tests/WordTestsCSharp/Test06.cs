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
    public class Test06 : ITestPackage
    {
        bool _beforeCloseCalled;
        bool _newDocumentCalled;

        #region TestPackage Member

        public string Name
        {
            get { return "Test06"; }
        }

        public string Description
        {
            get { return "Using events."; }
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

                application.NewDocumentEvent += new NetOffice.WordApi.Application_NewDocumentEventHandler(wordApplication_NewDocumentEvent);
                application.DocumentBeforeCloseEvent += new NetOffice.WordApi.Application_DocumentBeforeCloseEventHandler(wordApplication_DocumentBeforeCloseEvent);

                // add new document and close
                Word.Document document = application.Documents.Add();
                document.Close(false);

                if(_beforeCloseCalled && _newDocumentCalled)
                    return new TestResult(true, DateTime.Now.Subtract(startTime), "", null, "");
                else
                    return new TestResult(false, DateTime.Now.Subtract(startTime), string.Format("DocumentBeforeClose:{0}, NewDocument:{1}", _beforeCloseCalled, _newDocumentCalled), null, "");
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
         
        void wordApplication_DocumentBeforeCloseEvent(NetOffice.WordApi.Document Doc, ref bool Cancel)
        {
            _beforeCloseCalled = true;
            Doc.Dispose();
        }

        void wordApplication_NewDocumentEvent(NetOffice.WordApi.Document Doc)
        {
            _newDocumentCalled = true;
            Doc.Dispose();
        }
    }
}
