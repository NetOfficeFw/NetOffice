using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using System.Reflection;

namespace AccessingWordBasic
{
    class Program
    {
        static void Main(string[] args)
        {
            Word.Application application = null;
            try
            {
                application = new Word.ApplicationClass();
                application.DisplayAlerts = WdAlertLevel.wdAlertsNone;

                object basic = application.WordBasic;
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
                if (null != application)
                {
                    application.Quit();
                    application.Dispose();
                }
            }

            Console.ReadKey();
        }
    }
}
