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
    public class Tutorial04 : ITutorial
    {
        #region ITutorial

        public void Run()
        {
            // this example shows you how i still can recieve events from an disposed proxy.
            // you have to use th Dispose oder DisposeChildInstances method with a parameter.

            // start application
            Excel.Application application = new Excel.Application();
            application.DisplayAlerts = false;

            // create new Workbook & attach close event trigger
            Excel.Workbook book = application.Workbooks.Add();
            book.BeforeCloseEvent += new Excel.Workbook_BeforeCloseEventHandler(book_BeforeCloseEvent);

            // we dispose the instance. the parameter false signals to api dont release the event listener
            // set parameter to true and the event listener will stopped and you dont get events for the instance
            // the DisposeChildInstances() method has the same method overload
            book.Close();
            book.Dispose(false);

            application.Quit();
            application.Dispose();
            
            // the application object is ouer root object
            // dispose them release himself and any childs of application, in this case workbooks and workbook
            // the excel instance are now removed from process list

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
            get { return HostApplication.LCID == 1033 ? "http://netoffice.codeplex.com/wikipage?title=Tutorial04_EN_CS" : "http://netoffice.codeplex.com/wikipage?title=Tutorial04_DE_CS"; }
        }

        public string Caption
        {
            get { return "Tutorial04"; }
        }

        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Using Dispose with event exporting Objects" : "Verwenden von Dispose mit Objekten die Ereignisse auslösen"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion

        #region Properties

        internal IHost HostApplication { get; private set; }

        #endregion

        #region Trigger

        private void book_BeforeCloseEvent(ref bool Cancel)
        {

        }

        #endregion
    }
}
