using System;
using System.Collections.Generic;
using System.Text;

using LateBindingApi.Core;
using Excel = NetOffice.ExcelApi;

namespace UseTheDebugConsole
{
    class Program
    {
        static void Main(string[] args)
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
                Console.WriteLine("An error is occured. NetOffice DebugConsole content below:");

                foreach (string item in DebugConsole.Messages)
                    Console.WriteLine(item);

                Console.ReadKey();
            }
            finally
            {
                // quit and dispose
                application.Quit();
                application.Dispose();
            }

        }
    }
}
