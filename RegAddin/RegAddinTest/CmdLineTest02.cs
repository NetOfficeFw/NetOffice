using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using RegAddin;

namespace RegAddinTest
{
    internal class CmdLineTest02
    {     
        /// <summary>
        /// Try to give an unkown option
        /// </summary>
        /// <returns></returns>
        internal static bool Test1()
        {
            CommandLineValidator cmdValidator = new CommandLineValidator();
            string[] args = new string[] { " /foo " };

            try
            {
                cmdValidator.ValidateCommandLineArguments(ref args);
                return false;
            }
            catch(RegAddinException)
            {
                return true;
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Try to give an illegal combination of options
        /// </summary>
        /// <returns></returns>
        internal static bool Test2()
        {
            CommandLineValidator cmdValidator = new CommandLineValidator();
            string[] args = new string[] { "C:\\MyFile.dll", "/reg", "/unregfile" };

            try
            {
                cmdValidator.ValidateCommandLineArguments(ref args);
                return false;
            }
            catch (RegAddinException)
            {
                return true;
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Try to set /reg with invalid arguments
        /// </summary>
        /// <returns></returns>
        internal static bool Test3()
        {
            CommandLineValidator cmdValidator = new CommandLineValidator();
            string[] args = new string[] { "C:\\MyFile.dll", "/reg:foo" };

            try
            {
                cmdValidator.ValidateCommandLineArguments(ref args);
                return false;
            }
            catch (RegAddinException)
            {
                return true;
            }
            catch (Exception)
            {
                throw;
            }
        }


        /// <summary>
        /// Try to set /unregfile with invalid arguments
        /// </summary>
        /// <returns></returns>
        internal static bool Test4()
        {
            CommandLineValidator cmdValidator = new CommandLineValidator();
            string[] args = new string[] { "C:\\MyFile.dll", "/unregfile:C:\\MyFile.reg" };

            try
            {
                cmdValidator.ValidateCommandLineArguments(ref args);
                return false;
            }
            catch (RegAddinException)
            {
                return true;
            }
            catch (Exception)
            {
                throw;
            }
        }


        /// <summary>
        /// Just missing assembly on filesystem
        /// </summary>
        /// <returns></returns>
        internal static bool Test5()
        {
            CommandLineValidator cmdValidator = new CommandLineValidator();
            string[] args = new string[] { "C:\\MyFile.dll", "/reg" };

            try
            {
                cmdValidator.ValidateCommandLineArguments(ref args);
                return false;
            }
            catch (RegAddinException)
            {
                return true;
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
