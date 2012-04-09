using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.Text;
using ExampleBase;

using LateBindingApi.Core;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;

namespace MiscExamplesCS4
{
    /*
            *  in some situations of version independent developement, its necessary to check for
            *  the support of a specific entity at runtime. for this reason any object in NetOffice
            *  has the following method:
            *
            *  bool EntityIsAvailable(string name);
            *  bool EntityIsAvailable(string name, SupportEntityType searchType);
            *  
            *  this example shows you how to use them.
    */
    class Example02 : IExample
    {
        IHost _hostApplication;

        #region IExample Member

        public void RunExample()
        {
            // Initialize NetOffice
            LateBindingApi.Core.Factory.Initialize();

            // create excel instance
            Excel.Application application = new NetOffice.ExcelApi.Application();


            // ask the application object for Quit method support
            bool supportQuitMethod = application.EntityIsAvailable("Quit");

            // ask the application object for Visible property support
            bool supportVisbibleProperty = application.EntityIsAvailable("Visible");

            // ask the application object for SmartArtColors property support (only available in Excel 2010)
            bool supportSmartArtColorsProperty = application.EntityIsAvailable("SmartArtColors");

            // ask the application object for XYZ property or method support (not exists of course)
            bool supportTestXYZProperty = application.EntityIsAvailable("TestXYZ");

            // print result
            string messageBoxContent = "";
            messageBoxContent += string.Format("Your installed Excel Version supports the Quit Method: {0}{1}", supportQuitMethod, Environment.NewLine);
            messageBoxContent += string.Format("Your installed Excel Version supports the Visible Property: {0}{1}", supportVisbibleProperty, Environment.NewLine);
            messageBoxContent += string.Format("Your installed Excel Version supports the SmartArtColors Property: {0}{1}", supportSmartArtColorsProperty, Environment.NewLine);
            messageBoxContent += string.Format("Your installed Excel Version supports the TestXYZ Property: {0}{1}", supportTestXYZProperty, Environment.NewLine);
            MessageBox.Show(messageBoxContent, "EntityIsAvailable Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
             
            // quit and dispose
            application.Quit();
            application.Dispose();
        }

        public void Connect(IHost hostApplication)
        {
            _hostApplication = hostApplication;
        }

        public string Caption
        {
            get { return _hostApplication.LCID == 1033 ? "Example02" : "Beispiel02"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Check entity support at runtime" : "Zur Laufzeit prüfen ob eine Methode oder Property unterstützt wird"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion
    }
}
