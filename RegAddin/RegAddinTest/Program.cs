using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddinTest
{
    internal class Program
    {
        internal static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("RegAddinTest - Begin Tests");

                //if (!CmdLineTest01.Test1())
                //    Console.WriteLine("CmdLineTest01.Test1 failed.");
                //if (!CmdLineTest01.Test2())
                //    Console.WriteLine("CmdLineTest01.Test2 failed.");

                //if (!CmdLineTest02.Test1())
                //    Console.WriteLine("CmdLineTest02.Test1 failed.");
                //if (!CmdLineTest02.Test2())
                //    Console.WriteLine("CmdLineTest02.Test2 failed.");
                //if (!CmdLineTest02.Test3())
                //    Console.WriteLine("CmdLineTest02.Test3 failed.");
                //if (!CmdLineTest02.Test4())
                //    Console.WriteLine("CmdLineTest02.Test4 failed.");
                //if (!CmdLineTest02.Test5())
                //    Console.WriteLine("CmdLineTest02.Test5 failed.");

                //if (!CmdLineTest03.Test1())
                //    Console.WriteLine("CmdLineTest03.Test1 failed.");
                //if (!CmdLineTest03.Test2())
                //    Console.WriteLine("CmdLineTest03.Test2 failed.");
                //if (!CmdLineTest03.Test3())
                //    Console.WriteLine("CmdLineTest03.Test3 failed.");


                if (!CmdLineTest04.Test1())
                    Console.WriteLine("CmdLineTest04.Test1 failed.");
                //if (!CmdLineTest04.Test2())
                //    Console.WriteLine("CmdLineTest04.Test4 failed.");

                //if (!CmdLineTest04.Test3())
                //    Console.WriteLine("CmdLineTest04.Test3 failed.");
                //if (!CmdLineTest04.Test4())
                //    Console.WriteLine("CmdLineTest04.Test4 failed.");

                //if (!CmdLineTest04.Test5())
                //    Console.WriteLine("CmdLineTest04.Test5 failed.");

                Console.WriteLine("RegAddinTest - Tests passed");
                Console.ReadKey();
            }
            catch (Exception exception)
            {
                Console.WriteLine("RegAddinTest - Unexpected Error");
                Console.WriteLine(exception.ToString());
                Console.ReadKey();
            }
        }
    }
}
