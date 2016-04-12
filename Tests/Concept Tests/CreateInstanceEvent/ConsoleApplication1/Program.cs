using System;
using System.Collections.Generic;
using System.Text;
using NetOffice;
using Excel = NetOffice.ExcelApi;

namespace ConsoleApplication1
{
    class Program
    {
        public static void Main(string[] args)
        {
            Excel.Application app = null;
            try
            {
                System.Console.WriteLine("NetOffice CreateInstance Event Concept Test\r\n");

                // Use this.Factory.CreateInstance instead in NetOffice Tools COMAddin
                NetOffice.Core.Default.CreateInstance += new Core.OnCreateInstanceEventHandler(Default_CreateInstance);
                //Disable NO console for clean view/output
                NetOffice.Core.Default.Console.Mode = DebugConsoleMode.None;

                app = new Excel.Application();
                Excel.Workbook book = app.Workbooks.Add();
                MyCustomWorksheet sheet = book.Sheets[1] as MyCustomWorksheet;
                sheet.PrintName();

                Console.WriteLine("\r\nTest passed");
                Console.ReadKey();
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                Console.ReadKey();
            }
            finally
            {
                if (null != app)
                {
                    app.Quit();
                    app.Dispose();
                    app = null;
                }
            }
        }

        private static void Default_CreateInstance(Core sender, Core.OnCreateInstanceEventArgs args)
        {
            // we replace all Worksheet instances with MyCustomWorksheet

            Excel.Worksheet sheet =  args.Instance as Excel.Worksheet;
            if (null != sheet)
                args.Replace = typeof(MyCustomWorksheet);

            // keep in your mind: args.Instance.DisposeChildInstances() is called after this event trigger so dont handle your business logic here
        }
    }

    // A custom worksheet 
    public class MyCustomWorksheet : Excel.Worksheet
    {
        // some ctors we need

		public MyCustomWorksheet(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

		public MyCustomWorksheet(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		public MyCustomWorksheet(Core factory, COMObject parentObject, object comProxy, Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		public MyCustomWorksheet(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
        public MyCustomWorksheet(COMObject replacedObject): base(replacedObject)
		{
			
		}

        // extend worksheet with a sample method
        public void PrintName()
        {
            System.Console.WriteLine("MyCustomWorksheet {0}", Name);
        }
    }
}
