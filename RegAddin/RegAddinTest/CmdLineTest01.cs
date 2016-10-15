using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using RegAddin;

namespace RegAddinTest
{
    internal class CmdLineTest01
    {
        /// <summary>
        /// Remove '/',whitespaces and change '?' alias to its internal real option name 'help'
        /// </summary>
        /// <returns></returns>
        internal static bool Test1()
        {
            CommandLineValidator cmdValidator = new CommandLineValidator();
            string[] args = new string[] { " /? " };
            cmdValidator.ValidateCommandLineArguments(ref args);
            if (args[0] == "help")
                return true;
            else
                return false;
        }

        /// <summary>
        /// Remove '-',whitespace and change '?' alias to its internal real option name 'help'
        /// </summary>
        /// <returns></returns>
        internal static bool Test2()
        {
            CommandLineValidator cmdValidator = new CommandLineValidator();
            string[] args = new string[] { " -? " };
            cmdValidator.ValidateCommandLineArguments(ref args);
            if (args[0] == "help")
                return true;
            else
                return false;
        }
    }
}
