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
    public class Tutorial01 : ITutorial
    {
        #region ITutorial

        public void Run()
        {
            //  NetOffice manages COM Proxies for you to avoid any kind of memory leaks
            //  and make sure your application instance removes from process list if you want.

            // start application
            Excel.Application application = new Excel.Application();
            application.DisplayAlerts = false;

            Excel.Workbook book = application.Workbooks.Add();
            /* 
            * now we have 2 new COM Proxies created.
            * 
            * the first proxy was created while accessing the Workbooks collection from application
            * the second proxy was created by the Add() method from Workbooks and stored now in book
            * with the application object we have 3 created proxies now. the workbooks proxy was created
            * about application and the book proxy was created about the workbooks.
            * NetOffice holds the proxies now in a list as follows:
            * 
            * Application
            *   + Workbooks
            *     + Workbook  
            * 
            * any object in NetOffice implements the IDisposible Interface.
            * use the Dispose() method to release an object. the method release all created child proxies too.
            */

            application.Quit();
            application.Dispose();
            /*
            * the application object is ouer root object
            * dispose them release himself and any childs of application, in this case workbooks and workbook
            * the excel instance are now removed from process list
            */

            HostApplication.ShowFinishDialog();
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
            get { return HostApplication.LCID == 1033 ? "http://netoffice.codeplex.com/wikipage?title=Tutorial01_EN_CS" : "http://netoffice.codeplex.com/wikipage?title=Tutorial01_DE_CS"; }
        }

        public string Caption
        {
            get { return "Tutorial01"; }
        }


        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Understand COM Proxy Management" : "COM Proxy Management verstehen"; }
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
