using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using RegAddin;

namespace RegAddinTest
{
    internal class CmdLineTest04
    {
        // unreg test mit delete keys by regaddin 
        internal static bool Test1()
        {
            string[] args = new string[3];
            args[0] = @"C:\Sebastian\NetOffice11\Source\ClientAddin\bin\Debug\ClientAddin.dll";
            args[1] = "/unreg:auto:on:on:true";
            RegAddin.Program.Main(args);
            return true;
        }

        // reg test mit create keys by regaddin 
        internal static bool Test2()
        {
            string[] args = new string[3];
            args[0] = @"C:\Sebastian\NetOffice11\Source\ClientAddin\bin\Debug\ClientAddin.dll";
            args[1] = "/reg:system:on:on";
            RegAddin.Program.Main(args);
            return true;
        }

        // unreg test mit delete keys by himself 
        internal static bool Test3()
        {
            string[] args = new string[3];
            args[0] = @"C:\Sebastian\NetOffice11\Source\ClientAddin\bin\Debug\ClientAddin.dll";
            args[1] = "/unreg:auto:on:off:true";
            RegAddin.Program.Main(args);
            return true;
        }

        // reg test mit create keys by himself 
        internal static bool Test4()
        {
            string[] args = new string[3];
            args[0] = @"C:\Sebastian\NetOffice11\Source\ClientAddin\bin\Debug\ClientAddin.dll";
            args[1] = "/reg:user:on:off";
            RegAddin.Program.Main(args);
            return true;
        }

        // regfile test mit create keys by regaddin 
        internal static bool Test5()
        {
            string[] args = new string[3];
            args[0] = @"C:\Sebastian\NetOffice11\Source\ClientAddin\bin\Debug\ClientAddin.dll";
            args[1] = "/regfile:user:on:on:" + @"C:\Sebastian\NetOffice11\Source\ClientAddin\bin\Debug\ClientAddin.reg";
            RegAddin.Program.Main(args);
            return true;
        }
    }
}
