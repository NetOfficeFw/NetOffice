using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = NetOffice.OutlookApi;

namespace CreateOutlookAppIfRunning
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Create 3x Outlook Application");

            Outlook.Application application1 = new Outlook.Application();
            Outlook.Application application2 = new Outlook.Application();
            Outlook.Application application3 = new Outlook.Application();
            
            Console.WriteLine("Done! Press any key");
            Console.ReadKey();

            application1.Dispose();
            application2.Dispose();
            application3.Dispose();

            Console.WriteLine("Finish! Press any key..");
            Console.ReadKey();
        }
    }
}
