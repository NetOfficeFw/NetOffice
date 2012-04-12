using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using TutorialsBase;

using LateBindingApi.Core;
using Excel = NetOffice.ExcelApi;

namespace TutorialsCS4
{
    public partial class Tutorial05 : ITutorial
    { 
        IHost _hostApplication;

        #region ITutorial Member

        public void Run()
        {
            // Initialize NetOffice
            LateBindingApi.Core.Factory.Initialize();

            // start application
            Excel.Application application = new Excel.Application();
            application.DisplayAlerts = false;

            // create new Workbook
            Excel.Workbook book = application.Workbooks.Add();

            // ActiveSheet is defined as unkown Proxy in Excel Type Library, it can have multiple times at runtime
            // but its always a COM Proxy, never a scalar type like bool or int. 
            // In VBA oder PIA its converted to object, in NetOffice its also represents as object
            object sheet = application.ActiveSheet;
            if (sheet is Excel.Worksheet)
            {
                Excel.Worksheet activeSheet = (Excel.Worksheet)sheet;
            }

            // all classes inherites from the common base type COMObject
            // you can use also:
            COMObject unkownSheet = application.ActiveSheet as COMObject;

            // 3 basic properties of COMObject
            object proxy = unkownSheet.UnderlyingObject;            // the real COM proxy, be carefull !
            string proxyClassName = unkownSheet.UnderlyingTypeName; // the class name of the COM proxy, for example "Worksheet"
            bool isDisposed = unkownSheet.IsDisposed;               // info about the object is already disposed

            application.Quit();
            application.Dispose();

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
            get { return _hostApplication.LCID == 1033 ? "http://netoffice.codeplex.com/wikipage?title=Tutorial05_EN_CS" : "http://netoffice.codeplex.com/wikipage?title=Tutorial05_DE_CS"; }

        }

        public string Caption
        {
            get { return "Tutorial05"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Understanding unkown Types" : "Richtiges verwenden von unbekannten Typen"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion
    }
}
