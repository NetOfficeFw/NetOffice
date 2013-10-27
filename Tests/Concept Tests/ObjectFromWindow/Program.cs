using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;
using Accessibility;

namespace ObjectFromWindow
{
    class Program
    {
        static void Main(string[] args)
        {
            WindowEnumerator enumerator = new WindowEnumerator("XLMAIN");
            DateTime startTime = DateTime.Now;
            IntPtr[] handles = enumerator.EnumerateWindows(1000);
            if (null != handles)
            {
                foreach (IntPtr item in handles)
                {
                    object proxy = ExcelApplicationWindow.GetApplicationProxyFromHandle(item);
                    if (null != proxy)
                    {
                        NetOffice.ExcelApi.Application application = new NetOffice.ExcelApi.Application(null, proxy);
                        Console.WriteLine("Excel.Application Hwnd:{0}", application.Hwnd);
                    }
                }
            }
            else
            {
                Console.WriteLine("Enumerate Windows failed because the timeout is reached.");
            }

            Console.WriteLine("Press any key..");
            Console.ReadKey();
        }
    }
}
