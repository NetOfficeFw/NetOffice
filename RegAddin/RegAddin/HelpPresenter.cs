using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin
{
    internal class HelpPresenter
    {
        internal static int _padRightCount = 32;

        internal void ShowHelp()
        {
            Console.WriteLine("{0} -- Commandline Reference --{0}", Environment.NewLine);
            Console.WriteLine("Name:Options".PadRight(_padRightCount) + "Description{0}", Environment.NewLine);

            foreach (var item in new Commands())
                Console.WriteLine((item.Name + item.ArgumentsSyntax).PadRight(_padRightCount) + item.HelpText);

            Console.WriteLine("{0} -- Commandline Examples --{0}", Environment.NewLine);

            Console.WriteLine(" Example 1 - Register an addin to the system{0}", Environment.NewLine);
            Console.WriteLine(" \tRegAddin.exe \"C:\\MyFiles\\MyAddin.exe\" -reg{0}", Environment.NewLine);

            Console.WriteLine(" Example 2 - Register an addin to the current user{0}", Environment.NewLine);
            Console.WriteLine(" \tRegAddin.exe \"C:\\MyFiles\\MyAddin.exe\" -reg:local{0}", Environment.NewLine);

            Console.WriteLine(" Example 3 - Unregister an addin to the current user{0}", Environment.NewLine);
            Console.WriteLine(" \tRegAddin.exe \"C:\\MyFiles\\MyAddin.exe\" -unreg:local{0}", Environment.NewLine);

            Console.WriteLine("{0} -- See http://netoffice.codeplex for further documentation --{0}", Environment.NewLine);
        }      
    }
}
