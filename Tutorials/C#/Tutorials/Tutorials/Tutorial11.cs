using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using TutorialsBase;

namespace TutorialsCS4
{
    public partial class Tutorial11 : ITutorial
    {
        IHost _hostApplication;

        #region ITutorial Member

        public void Run()
        {
            string message = _hostApplication.LCID == 1033 ? "This tutorial doens't contain example code" : "Dieses Tutorial enthält keinen Beispielcode";
            MessageBox.Show(message, "Tutorial11", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void Connect(IHost hostApplication)
        {
            _hostApplication = hostApplication;
        }

        public void Disconnect()
        {

        }

        public void ChangeLanguage(int lcid)
        {

        }

        public string Uri
        {
            get { return _hostApplication.LCID == 1033 ? "http://netoffice.codeplex.com/wikipage?title=Tutorial11_EN_CS" : "http://netoffice.codeplex.com/wikipage?title=Tutorial11_DE_CS"; }
        }

        public string Caption
        {
            get { return "Tutorial11"; }
        }


        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Addin Deployment" : "Addins auf anderen System installieren"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion
    }
}
