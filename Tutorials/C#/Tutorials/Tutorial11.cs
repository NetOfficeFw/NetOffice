using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using TutorialsBase;

using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace TutorialsCS4
{
    public class Tutorial11 : ITutorial
    {
        #region ITutorial

        public void Run()
        {
            string message = HostApplication.LCID == 1033 ? "This tutorial doens't contain example code" : "Dieses Tutorial enthält keinen Beispielcode";
            MessageBox.Show(message, "Tutorial11", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public void Disconnect()
        {

        }

        public void ChangeLanguage(int lcid)
        {

        }

        public string Uri
        {
            get { return HostApplication.LCID == 1033 ? "http://netoffice.codeplex.com/wikipage?title=Tutorial11_EN_CS" : "http://netoffice.codeplex.com/wikipage?title=Tutorial11_DE_CS"; }
        }

        public string Caption
        {
            get { return "Tutorial11"; }
        }


        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Addin Deployment" : "Addins auf anderen System installieren"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion

        #region Properties

        internal IHost HostApplication { get; private set; }

        #endregion
    }
}
