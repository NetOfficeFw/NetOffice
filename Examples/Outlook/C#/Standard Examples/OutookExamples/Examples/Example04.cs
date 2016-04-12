using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
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
        #region IExample

        public void RunExample()
        {
            // start outlook
            Outlook.Application outlookApplication = new Outlook.Application();

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
            outlookApplication.Quit();
            outlookApplication.Dispose();

            HostApplication.ShowFinishDialog("Done!", null);
        }

        public string Caption
        {
            get { return HostApplication.LCID == 1033 ? "Example04" : "Beispiel04"; }
        }

        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Send and Recieve" : "Senden und empfangen"; }
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public System.Windows.Forms.UserControl Panel
        {
            get { return null; }
        }

        #endregion

        #region Properties

        internal IHost HostApplication { get; private set; }

        #endregion
    }
}
