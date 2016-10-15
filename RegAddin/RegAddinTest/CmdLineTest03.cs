using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using RegAddin;

namespace RegAddinTest
{
    internal class CmdLineTest03
    {
        /// <summary>
        /// Just call help
        /// </summary>
        /// <returns></returns>
        internal static bool Test1()
        {
            CommandLineSettingsTransformer cmdTransformer = new CommandLineSettingsTransformer();            
            string[] args = new string[] { "?" };
            cmdTransformer.ProceedCommandLineArguments(args);
            return SingletonSettings.Mode == SingletonSettings.ApplicationMode.Help;
        }

        /// <summary>
        /// Regfile operation
        /// </summary>
        /// <returns></returns>
        internal static bool Test2()
        {
            CommandLineSettingsTransformer cmdTransformer = new CommandLineSettingsTransformer();
            string[] args = new string[] { "C:\\MyFile.dll", "regfile:system:C:\\MyFile.reg", "codebase", "alert:on" };
            cmdTransformer.ProceedCommandLineArguments(args);
            if (SingletonSettings.Mode == SingletonSettings.ApplicationMode.RegFile &&
                SingletonSettings.AssemblyPath == "C:\\MyFile.dll" &&
                SingletonSettings.RegFilePath == "C:\\MyFile.reg" &&
                SingletonSettings.Codebase == true &&
                SingletonSettings.Alert == SingletonSettings.AlertMode.On)
                return true;
            else
                return false;
        }

        /// <summary>
        /// Unreg operation
        /// </summary>
        /// <returns></returns>
        internal static bool Test3()
        {
            CommandLineSettingsTransformer cmdTransformer = new CommandLineSettingsTransformer();
            string[] args = new string[] { "C:\\MyFile.dll", "unreg:user:false", "codebase", "alert" };
            cmdTransformer.ProceedCommandLineArguments(args);
            if (SingletonSettings.Mode == SingletonSettings.ApplicationMode.Unregister &&
                SingletonSettings.RegMode == SingletonSettings.RegisterMode.User &&
                SingletonSettings.AssemblyPath == "C:\\MyFile.dll" &&
                SingletonSettings.Codebase == true &&
                SingletonSettings.Alert == SingletonSettings.AlertMode.Error)
                return true;
            else
                return false;
        }
    }
}
