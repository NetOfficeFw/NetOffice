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
    public class Test03 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test03"; }
        }

        public string Description
        {
            get { return "Send a mail."; }
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

                // Create a new MailItem.
                Outlook.MailItem mailItem = application.CreateItem(OlItemType.olMailItem) as Outlook.MailItem;

                // prepare item and send
                mailItem.Recipients.Add("public.sebastian@web.de");
                mailItem.Subject = "NetOffice Test Mail";
                mailItem.Body = "This is a NetOffice test mail from the MainTests.(C#)";
                mailItem.Send();

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
                    application.Quit();
                    application.Dispose();
                }
            }
        }

        #endregion
    }
}
