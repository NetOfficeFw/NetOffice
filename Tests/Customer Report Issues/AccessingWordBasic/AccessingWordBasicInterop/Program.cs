using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Runtime.InteropServices;

namespace AccessingWordBasicInterop
{
    class Program
    {
        static void Main(string[] args)
        {
            Word.Application application = null;
            object basic = null;
            try
            {
                application = new Word.Application();
                application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                basic = application.WordBasic;
                object[] argValues = { 1 };
                basic.GetType().InvokeMember("DisableAutoMacros", BindingFlags.InvokeMethod, null, basic, argValues, null, null, null);
                Console.WriteLine("Fine");
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }
            finally
            {
                if (null != basic)
                {
                    Marshal.ReleaseComObject(basic);
                }

                if (null != application)
                {
                    application.Quit();
                    Marshal.ReleaseComObject(application);
                }
            }

            Console.ReadKey();
        }
    }
}
