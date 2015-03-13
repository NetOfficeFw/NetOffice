using System;
using System.Collections.Generic;
using System.Text;
using NetOffice.OutlookSecurity;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;

namespace SuppressOutlookSecurity
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Concept Test - SuppressOutlookSecurity");
            Suppress.Enabled = true;
            Suppress.OnAction += new Suppress.SecurityDialogAction(Suppress_OnAction);
            Suppress.OnError += new Suppress.ErrorOccuredEventHandler(Suppress_OnError);
            Outlook.Application application = null;
            try
            {
                application = new Outlook.Application();
                SendMail(application);
                Console.WriteLine("Press any key...");
                Console.ReadKey();
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.ToString());
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

        private static void SendMail(Outlook.Application application)
        {
            Outlook.MailItem mailItem = application.CreateItem(OlItemType.olMailItem) as Outlook.MailItem;
            mailItem.Recipients.Add("public.sebastian@web.de");
            mailItem.Subject = "Concept Test - SuppressOutlookSecurity";
            mailItem.Body = "This is a test mail from NetOffice concept test.";
            mailItem.Send();
        }


        private static void Suppress_OnError(Exception exception)
        {
            Console.WriteLine("Supress_OnError:{0}{1}", Environment.NewLine, exception);
        }

        private static void Suppress_OnAction(SecurityDialog dialog, SecurityDialogCheckBox targetBox, SecurityDialogLeftButton targetButton)
        {
            Console.WriteLine("Suppress_OnAction:{0}{1}{2}", Environment.NewLine, dialog, Environment.NewLine, targetButton);
        }
    }
}
