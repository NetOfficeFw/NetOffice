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
    public class Test04 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test04"; }
        }

        public string Description
        {
            get { return "Perform send and recieve."; }
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

                if (application.Session.EntityIsAvailable("SendAndReceive"))
                {
                    application.Session.SendAndReceive(false);
                    // give few seconds to outlook or may its failed to quit because its busy - depending on how many mails comes in
                    System.Threading.Thread.Sleep(3000);
                }
                else
                    return new TestResult(false, DateTime.Now.Subtract(startTime), "SendAndReceive is not supported from this Outlook Version.", null, "");

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
