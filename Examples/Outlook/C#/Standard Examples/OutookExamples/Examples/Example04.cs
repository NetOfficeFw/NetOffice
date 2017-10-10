using System;
using ExampleBase;
using NetOffice;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;

namespace OutlookExamplesCS4
{
    /// <summary>
    /// Example 4 - Send and recieve
    /// </summary>
    internal class Example04 : IExample
    {
        public void RunExample()
        {
            // start outlook by trying to access running application first
            Outlook.Application outlookApplication = new Outlook.Application(true);

            // SendAndReceive is supported from Outlook 2007 or higher
            // we check at runtime the feature is available
            if (outlookApplication.Session.EntityIsAvailable("SendAndReceive"))
            {
                // one simple call
                outlookApplication.Session.SendAndReceive(false);
            }
            else
            {
                HostApplication.ShowErrorDialog("This version of MS-Outlook doesnt support SendAndReceive.", null);
            }

            // close outlook and dispose
            if (!outlookApplication.FromProxyService)
                outlookApplication.Quit();
            outlookApplication.Dispose();

            HostApplication.ShowFinishDialog("Done!", null);
        }

        public string Caption
        {
            get { return "Example04"; }
        }

        public string Description
        {
            get { return "Send and Recieve"; }
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public System.Windows.Forms.UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}
