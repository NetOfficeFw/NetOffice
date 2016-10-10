using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin.Metrics
{
    internal class ConsolePresenter
    {
        internal void Show(Dictionary<string, bool> result)
        {
            Console.WriteLine("The following metric rules failed to validate:");

            foreach (var item in result)
            {
                if (item.Value == false)
                    Console.WriteLine(item.Key);
            }

            Console.WriteLine("");
        }
    }
}
