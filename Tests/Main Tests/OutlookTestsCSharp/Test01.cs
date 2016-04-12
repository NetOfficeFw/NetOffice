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
    public class Test01 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test01"; }
        }

        public string Description
        {
            get { return "Fetch inbox folder."; }
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

                // Get inbox 
                Outlook._NameSpace outlookNS = application.GetNamespace("MAPI");
                Outlook.MAPIFolder inboxFolder = outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
              
                Outlook._Items items = inboxFolder.Items;
                COMObject item = null;
                int i = 1;
                do
                {
                    if (null == item)
                        item = items.GetFirst() as COMObject;

                    // not every item is a mail item
                    Outlook.MailItem mailItem = item as Outlook.MailItem;
                    if (null != mailItem)
                        Console.WriteLine(mailItem.SenderName);
                    
                    if(null != item)
                        item.Dispose();

                    item = items.GetNext() as COMObject;
                    i++;
                } while (null != item);
              
                return new TestResult(true, DateTime.Now.Subtract(startTime), "", null, string.Format("{0} Inbox Items.", items.Count));
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
