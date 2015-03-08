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
    public class Tutorial10 : ITutorial
    {
        #region ITutorial

        public void Run()
        {
            // this property allows you to disable any events from Office applications
            bool enableEvents = NetOffice.Settings.Default.EnableEvents;

            // this property is the common threadculture for accessing Office.
            // default is en-us(1033)
            System.Globalization.CultureInfo threadCulture = NetOffice.Settings.Default.ThreadCulture;
             
            // this property allows to enable a COM Message filter
            // you may know the problem with RPC_Rejected exceptions(the host application is currently busy)
            // the message filter is waiting for you(shortly) and try the call again
            bool messageFilter = NetOffice.Settings.Default.MessageFilter.Enabled;
              
            // the safemode is a feature that checks automaticly at runtime the methods oder properties you use are
            // available in current office version. if it doesnt an EntityNotSupportedException was thrown
            // false by default
            bool safeMode = NetOffice.Settings.Default.EnableSafeMode;
            
            // get or set NetOffice logs essential system steps in the NetOffice DebugConsole
            bool debugOutput = NetOffice.Settings.Default.EnableDebugOutput;

            // this property allows you to enable NetOffice call Quit() for Application objects automaticly while Dispose()
            // false by default
            bool automaticQuit = NetOffice.Settings.Default.EnableAutomaticQuit;

            string message = string.Format("Events enabled:{0}{6}Thread:{1}{6}AutomaticQuit enabled:{2}{6}MessageFilter enabled:{3}{6}SafeMode enabled:{4}{6}DebugOutput enabled:{5}{6}", enableEvents, threadCulture.LCID, automaticQuit, messageFilter, safeMode, debugOutput, Environment.NewLine);
            MessageBox.Show(message, "Settings", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            get { return HostApplication.LCID == 1033 ? "http://netoffice.codeplex.com/wikipage?title=Tutorial10_EN_CS" : "http://netoffice.codeplex.com/wikipage?title=Tutorial10_DE_CS"; }
        }

        public string Caption
        {
            get { return "Tutorial10"; }
        }


        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "NetOffice Settings" : "Einstellungsmöglichkeiten für NetOffice"; }
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
