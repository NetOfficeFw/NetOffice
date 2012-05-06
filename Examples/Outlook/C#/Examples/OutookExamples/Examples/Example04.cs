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
    class Example04 : IExample
    {
        IHost _hostApplication;

        #region IExample Member

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
                _hostApplication.ShowErrorDialog("This version of MS-Outlook doesnt support SendAndReceive.", null);
            }

            // close outlook and dispose
            outlookApplication.Quit();
            outlookApplication.Dispose();

		 _hostApplication.ShowFinishDialog("Done!", null);
        }

        public string Caption
        {
            get { return _hostApplication.LCID == 1033 ? "Example04" : "Beispiel04"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Send and Recieve" : "Senden und empfangen"; }
        }

        public void Connect(IHost hostApplication)
        {
            _hostApplication = hostApplication;
        }

        public System.Windows.Forms.UserControl Panel
        {
            get { return null; }
        }

        #endregion
    }
}
