using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelPrototypeTest
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                new Test03().Run();
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }
            finally
            {
                Console.WriteLine("Press any key...");
                Console.ReadKey();
            }
        }
    }
}
