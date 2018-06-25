using System;
using NetOffice;

namespace ObjectFromWindow
{
    class Program
    {
        /// <summary>
        ///
        /// </summary>
        /// <param name="args"></param>
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
                        NetOffice.ExcelApi.Application application = COMObject.Create<NetOffice.ExcelApi.Application>(proxy);
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
