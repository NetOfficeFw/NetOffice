using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UnhookEvents
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                new Test01().Proceed();
            }
            catch (Exception exception)
            {
                Console.WriteLine("Fatal Error{1}{0}{1}Press any key to abort.", exception.Message, Environment.NewLine);
                Console.ReadKey();
            }
        }
    }
}
