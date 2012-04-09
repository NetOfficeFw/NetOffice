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
             *  NetOffice gives you as additional service a debug console.
             *  Essential NetOffice system steps and any occured exception related with
             *  your office application(or NetOffice himself maybe) are stored here. if an error occureed and you need help,
             *  please use the NetOffice discussion board: http://netoffice.codeplex.com/discussions
             *  describe your problem and post the content of the DebugConsole as below your message.
             *  the following infos are also helpful: operating system 32 or 64 bit, office version 32 or 64 bit, assembly runs as administrator or not
             *
             *  the following options are available
             * 
             *   ConsoleMode.None       = Console is deactivated (default)
             *   ConsoleMode.Console    = redirect all messages to System.Console
             *   ConsoleMode.MemoryList = keep all messages in memory. use DebugConsole.Messages and DebugConsole.ClearMessagesList() with these option
             *   ConsoleMode.LogFile    = writes all messages immediately to a file. you have to set DebugConsole.FileName before use
     */
    class Example01 : IExample
    {
        IHost _hostApplication;
        
        #region IExample Member

        public void RunExample()
        {
            // Initialize NetOffice
            LateBindingApi.Core.Factory.Initialize();

            Excel.Application application = null;
            try
            {
                // activate the DebugConsole. the default value is: ConsoleMode.None
                DebugConsole.Mode = ConsoleMode.MemoryList;

                // Initialize NetOffice
                LateBindingApi.Core.Factory.Initialize();

                // create excel instance
                application = new NetOffice.ExcelApi.Application();
                application.DisplayAlerts = false;

                // we open a non existing file to produce an error
                application.Workbooks.Open("z:\\NotExistingFile.exe");
            }
            catch (Exception)
            {
                string messages = null;
                foreach (string item in DebugConsole.Messages)
                    messages += item + Environment.NewLine;

               MessageBox.Show(messages, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // quit and dispose
                application.Quit();
                application.Dispose();
            }
        }

        public void Connect(IHost hostApplication)
        {
            _hostApplication = hostApplication;
        }

        public string Caption
        {
            get { return _hostApplication.LCID == 1033 ? "Example01" : "Beispiel01"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Use the NetOffice Debug Console" : "Benutzen der NetOffice Debug Console"; }
        }

        public UserControl Panel 
        {
            get { return null; }
        }

        #endregion
    }
}
