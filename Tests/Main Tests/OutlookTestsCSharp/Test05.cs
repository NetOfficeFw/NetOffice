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
    public class Test05 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test05"; }
        }

        public string Description
        {
            get { return "Fetch Contacts."; }
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

                // enum contacts 
                Outlook.MAPIFolder contactFolder = application.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
                for (int i = 1; i <= contactFolder.Items.Count; i++)
                {
                    Outlook.ContactItem contact = contactFolder.Items[i] as Outlook.ContactItem;
                    if (null != contact)
                        Console.WriteLine(contact.CompanyAndFullName);
                }

                return new TestResult(true, DateTime.Now.Subtract(startTime), "", null, string.Format("{0} ContactFolder Items.", contactFolder.Items.Count));
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
