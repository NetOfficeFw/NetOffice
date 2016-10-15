using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace RegAddin
{
    /// <summary>
    /// Transform commandline arguments in SingletonSettings. 
    /// Before using this class the CommandLineValidator must run because this transformation unit is not fail safe.
    /// </summary>
    internal class CommandLineSettingsTransformer
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="args"></param>
        internal void ProceedCommandLineArguments(string[] args)
        {
            foreach (string item in args)
            {
                Command command = GetTargetCommand(item, false);
                if (null == command)
                { 
                    SingletonSettings.AssemblyPath = item;
                    continue;
                }

                switch (command.Name)
                {
                    case "reg":
                        SingletonSettings.Mode = SingletonSettings.ApplicationMode.Register;
                        SingletonSettings.RegMode = GetRegisterModeOptionArgument(item, 0);
                        SingletonSettings.DoRegisterCall = GetRegisterCallOptionArgument(item, 1);
                        SingletonSettings.AddinReg = GetAddinRegOptionArgument(item, 2);
                        break;
                    case "unreg":
                        SingletonSettings.Mode = SingletonSettings.ApplicationMode.Unregister;
                        SingletonSettings.UnRegMode =  GetUnRegisterModeOptionArgument(item, 0);
                        SingletonSettings.DoRegisterCall = GetRegisterCallOptionArgument(item, 1);
                        SingletonSettings.AddinReg = GetAddinRegOptionArgument(item, 2);
                        SingletonSettings.SuspendMissingAssemblyErrorInUnregister = GetBoolOptionArgument(item, 3);
                        break;
                    case "regfile":
                        SingletonSettings.Mode = SingletonSettings.ApplicationMode.RegFile;
                        SingletonSettings.RegMode = GetRegisterModeOptionArgument(item, 0);
                        SingletonSettings.ExportCall = GetRegExportCallOptionArgument(item, 1);
                        SingletonSettings.AddinReg = GetAddinRegOptionArgument(item, 2);
                        SingletonSettings.RegFilePath = GetCommandOptionPathLastArgument(item, 3);
                        break;
                    case "help":
                        SingletonSettings.Mode = SingletonSettings.ApplicationMode.Help;
                        break;
                    case "alert":
                        SingletonSettings.Alert = GetAlertModeOptionArgument(item, 0);
                        break;
                    case "codebase":
                        SingletonSettings.Codebase = true;
                        break;
                    case "sign":
                        SingletonSettings.SignCheck = GetSignCheckOptionArgument(item, 0);
                        break;
                    case "diag":
                        SingletonSettings.Diagnostics = true;
                        break;
                    case "metrics":
                        SingletonSettings.Metrics = GetMetricsOptionArgument(item, 0);
                        break;
                    default:
                        break;
                }
            }
        }

        private bool GetBoolOptionArgument(string fullCommandOption, int argumentIndex)
        {
            int validatedIndex = argumentIndex + 1;
            string[] array = fullCommandOption.Split(new string[] { ":" }, StringSplitOptions.None);
            string argument = array[validatedIndex].ToLower();
            return bool.Parse(argument);
        }

        private T GetEnumOptionArgument<T>(string fullCommandOption, int argumentIndex)
        {
            int validatedIndex = argumentIndex + 1;
            string[] array = fullCommandOption.Split(new string[] { ":" }, StringSplitOptions.None);
            string argument = array[validatedIndex];
            argument = argument.Substring(0, 1).ToUpper() + argument.Substring(1).ToLower();
            return (T)Enum.Parse(typeof(T), argument);
        }

        private SingletonSettings.UnRegisterMode GetUnRegisterModeOptionArgument(string fullCommandOption, int argumentIndex)
        {
            return GetEnumOptionArgument<SingletonSettings.UnRegisterMode>(fullCommandOption, argumentIndex);
        }

        private SingletonSettings.RegExportCall GetRegExportCallOptionArgument(string fullCommandOption, int argumentIndex)
        {
            return GetEnumOptionArgument<SingletonSettings.RegExportCall>(fullCommandOption, argumentIndex);
        }

        private SingletonSettings.AddinRegMode GetAddinRegOptionArgument(string fullCommandOption, int argumentIndex)
        {
            return GetEnumOptionArgument<SingletonSettings.AddinRegMode>(fullCommandOption, argumentIndex);
        }

        private SingletonSettings.MetricsMode GetMetricsOptionArgument(string fullCommandOption, int argumentIndex)
        {
            return GetEnumOptionArgument<SingletonSettings.MetricsMode>(fullCommandOption, argumentIndex);
        }

        private SingletonSettings.SignCheckMode GetSignCheckOptionArgument(string fullCommandOption, int argumentIndex)
        {
            return GetEnumOptionArgument<SingletonSettings.SignCheckMode>(fullCommandOption, argumentIndex);
        }

        private SingletonSettings.RegisterCall GetRegisterCallOptionArgument(string fullCommandOption, int argumentIndex)
        {
            return GetEnumOptionArgument<SingletonSettings.RegisterCall>(fullCommandOption, argumentIndex);
        }

        private SingletonSettings.AlertMode GetAlertModeOptionArgument(string fullCommandOption, int argumentIndex)
        {
            return GetEnumOptionArgument<SingletonSettings.AlertMode>(fullCommandOption, argumentIndex);
        }

        private SingletonSettings.RegisterMode GetRegisterModeOptionArgument(string fullCommandOption, int argumentIndex)
        {
            return GetEnumOptionArgument<SingletonSettings.RegisterMode>(fullCommandOption, argumentIndex);
        }
        
        private string GetCommandOptionPathLastArgument(string fullCommandOption, int argumentIndex)
        {
            int validatedIndex = argumentIndex + 2;
            string[] array = fullCommandOption.Split(new string[] { ":" }, StringSplitOptions.None);
            string result = String.Empty;
            for (int i = validatedIndex; i < array.Length; i++)
                result += array[i] + ":";
            if (result.EndsWith(":", StringComparison.InvariantCultureIgnoreCase))
                result = result.Substring(0, result.Length - 1);
            result = array[array.Length - 2] + ":" + result;
            return result;
        }
        
        private string GetCommandOptionArgument(string fullCommandOption, int argumentIndex)
        {
            int validatedIndex = argumentIndex + 1;
            string[] array = fullCommandOption.Split(new string[] { ":" }, StringSplitOptions.None);
            return array[validatedIndex];
        }

        private Command GetTargetCommand(string fullCommandOption, bool throwException)
        {
            string targetName = fullCommandOption;
            if (targetName.IndexOf(":") > -1)
                targetName = targetName.Substring(0, targetName.IndexOf(":"));

            Commands commands = new Commands();
            foreach (Command item in commands)
            {
                if (targetName == item.Name)
                    return item;
                foreach (var alias in item.Alias)
                {
                    if (targetName == alias)
                        return item;
                }
            }

            if (throwException)
                throw new ArgumentOutOfRangeException("fullCommandOption");
            else
                return null;
        }
    }
}
