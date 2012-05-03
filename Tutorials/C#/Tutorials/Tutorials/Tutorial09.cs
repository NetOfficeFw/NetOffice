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
    public partial class Tutorial09 : ITutorial
    {
        IHost _hostApplication;

        #region ITutorial Member

        public void Run()
        {
            // In some situations you want use NetOffice with an existing proxy, its typical for COM Addins.
            // this examples show you how its possible

            // we create a native Excel proxy
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            object excelProxy = Activator.CreateInstance(excelType);
            
            // we create an Excel Application object with the proxy as parameter,
            // excel is now under control by NetOffice
            Excel.Application excelApplication = new Excel.Application(null, excelProxy);

            excelApplication.Quit();
            excelApplication.Dispose();

            _hostApplication.ShowFinishDialog();
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
            get { return _hostApplication.LCID == 1033 ? "http://netoffice.codeplex.com/wikipage?title=Tutorial09_EN_CS" : "http://netoffice.codeplex.com/wikipage?title=Tutorial09_DE_CS"; }
        }

        public string Caption
        {
            get { return "Tutorial09"; }
        }


        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Create a NetOffice Application with given COM Proxy" : "Eine NetOffice Application basierend auf einem COM Proxy erstellen"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion
    }
}
