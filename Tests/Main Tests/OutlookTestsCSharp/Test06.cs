using System;
using System.Collections.Generic;
using System.Text;
using Tests.Core;
using NetOffice;
using Office = NetOffice.OfficeApi;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;

namespace OutlookTestsCSharp
{
    public class Test06 : ITestPackage
    {
        bool _closeEventCalled;

        #region TestPackage Member

        public string Name
        {
            get { return "Test06"; }
        }

        public string Description
        {
            get { return "Test events."; }
        }

        public string OfficeProduct
        {
            get { return "Outlook"; }
        }

        public string Language
        {
            get { return "C#"; }
        }

        public TestResult DoTest()
        {
            Outlook.Application application = null;
            DateTime startTime = DateTime.Now;
            try
            {
                // start outlook
                application = new Outlook.Application();
                NetOffice.OutlookSecurity.Suppress.Enabled = true;

                Outlook.MailItem mailItem = application.CreateItem(OlItemType.olMailItem) as Outlook.MailItem;
                mailItem.CloseEvent += new NetOffice.OutlookApi.MailItem_CloseEventHandler(mailItem_CloseEvent);

                // BodyFormat is not available in Outlook 2000
                // we check at runtime is property is available
                if (mailItem.EntityIsAvailable("BodyFormat"))
                    mailItem.BodyFormat = OlBodyFormat.olFormatPlain;
                mailItem.Body = "NetOffice C# Test06" + DateTime.Now.ToLongTimeString();
                mailItem.Subject = "Test06";
                mailItem.Display();
                mailItem.Close(OlInspectorClose.olDiscard);

                if(_closeEventCalled)
                    return new TestResult(true, DateTime.Now.Subtract(startTime), "", null, "");
                else
                    return new TestResult(false, DateTime.Now.Subtract(startTime), "CloseEvent not triggered.", null, "");
            }
            catch (Exception exception)
            {
                return new TestResult(false, DateTime.Now.Subtract(startTime), exception.Message, exception, "");
            }
            finally
            {
                if (null != application)
                {
                    application.Quit();
                    application.Dispose();
                }
            }
        }

        #endregion

        void mailItem_CloseEvent(ref bool Cancel)
        {
            _closeEventCalled = true;
        }
    }
}
