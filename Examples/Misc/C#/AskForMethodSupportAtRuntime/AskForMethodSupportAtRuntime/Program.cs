using System;
using System.Collections.Generic;
using System.Text;
using Excel = NetOffice.ExcelApi;

namespace AskForMethodSupportAtRuntime
{
    class Program
    {
        static void Main(string[] args)
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
            Console.WriteLine("Your installed Excel Version supports the Quit Method: {0}", supportQuitMethod);
            Console.WriteLine("Your installed Excel Version supports the Visbible Property: {0}", supportVisbibleProperty);
            Console.WriteLine("Your installed Excel Version supports the SmartArtColors Property: {0}", supportSmartArtColorsProperty);
            Console.WriteLine("Your installed Excel Version supports the TestXYZ Property: {0}", supportTestXYZProperty);
            Console.ReadKey();

            // quit and dispose
            application.Quit();
            application.Dispose();
        }
    }
}
