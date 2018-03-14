using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointCustomDocProperties
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                int countToTry = 10;
                Console.WriteLine("Test - Adding custom document property {0} times.", countToTry);
                for (int i = 1; i <= countToTry; i++)
                {
                    new Test01().Proceed(i);
                }
                Console.WriteLine("Test passed.");
            }
            catch (Exception exception)
            {
                Console.WriteLine("---Unexcepted Error.---");
                Console.WriteLine(exception.Message);
            }
            finally
            {
                Console.WriteLine("Press any key...");
                Console.ReadKey();
            }
        }
    }
}
