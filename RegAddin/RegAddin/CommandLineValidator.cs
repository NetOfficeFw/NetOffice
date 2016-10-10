using System;
using System.IO;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;
using System.Text;

namespace RegAddin
{
    /// <summary>
    /// Handle the syntax check/validation for all commandline arguments. This must be called at first.
    /// The validator want create a validated command line arguments copy and try to fix some possible minor errors at this time.
    /// Moreover the validator add missing optional arguments in the argument string(s) with its default value.
    /// </summary>
    internal class CommandLineValidator
    {
        #region Internal Methods

        /// <summary>
        /// Validate arguments from command line
        /// </summary>
        /// <param name="args">command line arguments - the method want create the arguments array</param>
        internal void ValidateCommandLineArguments(ref string[] args)
        {
            if (null == args || args.Length == 0)
                throw new RegAddinException("NoArguments");

            List<string> list = CreateListCopy(args);
            RemoveEmptySpaceFromList(list);
            RemovePreQuotesFromList(list);
            ReplaceAliasToOriginName(list);
            HandleMostSimpleCase(list);
            CheckCommandOptionsIsKnown(list);
            CheckCommandCombination(list);
            CheckAssemblyExistsIfNecessary(list);
            CheckOptionSyntax(list);          
            AddMissingDefaultArguments(list);
            args = CreateValidateArrayFromList(list);
        }

        #endregion

        #region Private Methods

        private void HandleMostSimpleCase(List<string> list)
        {
            if (list.Count == 1 && File.Exists(list[0]))
                list.Add("reg");          
        }

        private void ReplaceAliasToOriginName(List<string> list)
        {
            for (int i = 0; i < list.Count; i++)
            {
                string item = list[i];
                Command itemCommand = GetTargetCommand(item, false);
                if (null == itemCommand)
                    continue;
                 
                if (item.IndexOf(":") > -1)
                {
                    string targetName = item.Substring(0, item.IndexOf(":"));                  
                    if (!itemCommand.Name.Equals(targetName, StringComparison.InvariantCultureIgnoreCase))
                    {
                        item = itemCommand.Name + item.Substring(item.IndexOf(":"));
                        list[i] = item;
                    }
                }
                else
                {
                    if (!itemCommand.Name.Equals(item, StringComparison.InvariantCultureIgnoreCase))
                    {
                        item = itemCommand.Name;
                        list[i] = item;
                    }
                }

            }
        }

        private void AddMissingDefaultArguments(List<string> list)
        {
            for (int i = 0; i < list.Count; i++)
            {
                string item = list[i];
                Command itemCommand = GetTargetCommand(item, false);
                if (null == itemCommand || null == itemCommand.Description)
                    continue;

                int index = 0;
                foreach (CommandOptionDescription description in itemCommand.Description)
                {
                    string commandOptionArgument = TryGetCommandOptionArgument(item, index);
                    if (true == description.IsOptional && null != description.DefaultValueIfOptional && null == commandOptionArgument)
                    { 
                        item += ":" + description.DefaultValueIfOptional;
                        list[i] = item;
                    }
                    index++;
                }        
            }
        }

        private string TryGetCommandOptionArgument(string fullCommandOption, int argumentIndex)
        {
            int validatedIndex = argumentIndex + 1;
            string[] array = fullCommandOption.Split(new string[] { ":" }, StringSplitOptions.None);
            if (array.Length > validatedIndex)
                return array[validatedIndex];
            else
                return null;
        } 

        private void CheckOptionSyntax(List<string> list)
        {
            foreach (string item in list)
            {
                Command command = GetTargetCommand(item, false);
                if (null != command && null != command.Description)
                {
                    int index = 0;
                    foreach (CommandOptionDescription description in command.Description)
                    {
                        string argument = TryGetCommandOptionArgument(item, index);
                        if (null == argument && false == description.IsOptional)
                            throw new RegAddinException("MissingArgument");
                        else if (null != argument && description.IsEnumOption && false == description.IsValueValid(argument))
                            throw new RegAddinException("InvalidArgumentValue");
                        index++;
                    }
                }
            }
        }
          
        private void CheckAssemblyExistsIfNecessary(List<string> list)
        {
            bool isNecessary = false;
            Commands commands = new Commands();
            foreach (string item in list)
            {
                foreach (Command command in commands)
                {
                    if (item.StartsWith(command.Name, StringComparison.InvariantCultureIgnoreCase) && command.NeedAssemblyPath == AssemblyRequired.Yes)
                    { 
                        isNecessary = true;
                        break;
                    }
                }
            }

            if (isNecessary)
            {                
                string fullFilePath = list[0];
                if (!File.Exists(fullFilePath))
                    throw new RegAddinException("AssemblyNotFound");
            }
        }

        private void CheckCommandCombination(List<string> list)
        {
            Commands commands = new Commands();
            CommandsSyntax commandsSyntax = commands.Syntax;
            string rootCommandArgument = null;
            CommandSyntax rootCommandSyntax = null;

            int rootCommandsCount = 0;
            foreach (string item in list)
            {
                foreach (CommandSyntax command in commandsSyntax)
                {
                    if (item.StartsWith(command.Underlying.Name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        rootCommandSyntax = command;
                        rootCommandArgument = item;
                        rootCommandsCount++;
                    }
                }                
            }

            if (rootCommandsCount < 1)
            {
                foreach (var item in list)
                {
                    System.Windows.Forms.MessageBox.Show(item);
                }
                throw new RegAddinException("MissingCommandOption");
            }
            if (rootCommandsCount > 2)
                throw new RegAddinException("AmbiguousArguments");

            foreach (string item in list)
            {
                if (item == rootCommandArgument)
                    continue;
                Command targetCommand = GetTargetCommand(item, false);
                if (null == targetCommand)
                    continue;

                bool foundInAllowedSyntax = rootCommandSyntax.Items.Count() == 0;
                foreach (CommandSyntax syntax in rootCommandSyntax.Items)
                {
                    if (item.StartsWith(syntax.Underlying.Name, StringComparison.InvariantCultureIgnoreCase))
                    {
                        foundInAllowedSyntax = true;
                        break;
                    }
                }
                if (!foundInAllowedSyntax)
                    throw new RegAddinException("InvalidCombination");
            }           
        }

        private bool CheckCommandOptionIsKnown(string item, Commands commands)
        {
            bool knownCommand = false;
            foreach (var command in commands)
            {
                if (item.StartsWith(command.Name, StringComparison.InvariantCultureIgnoreCase))
                {
                    knownCommand = true;
                    break;
                }
                bool matchAlias = false;
                foreach (var alias in command.Alias)
                {
                    if (item.StartsWith(alias, StringComparison.InvariantCultureIgnoreCase))
                    {
                        matchAlias = true;
                        break;
                    }
                }
                if (matchAlias)
                {
                    knownCommand = true;
                    break;
                }
            }

            return knownCommand;          
        }

        private void CheckCommandOptionsIsKnown(List<string> list)
        {
            Commands commands = new Commands();
            
            for (int i = 0; i < list.Count; i++)
            {
                string item = list[i];
                if (i == 0 )
                {
                    //if (!CheckCommandOptionIsKnown(item, commands))
                    //{
                    //    // this is a dirty solution right now - we need the check the list doesnt contains commands there want an assembly path
                    //    if (false == list.Any(e => e.StartsWith("reg", StringComparison.InvariantCultureIgnoreCase)) && false == list.Any(e => e.StartsWith("unreg", StringComparison.InvariantCultureIgnoreCase)) &&
                    //       false == list.Any(e => e.StartsWith("regfile", StringComparison.InvariantCultureIgnoreCase)))
                    //        throw new RegAddinException("InvalidArguments");
                    //}
                }
                else
                {
                    if (!CheckCommandOptionIsKnown(item, commands))
                    {                         
                        throw new RegAddinException("UnkownArguments");
                    }
                }
            }
        }

        private string[] CreateValidateArrayFromList(List<string> list)
        {
            string[] args = new string[list.Count];
            for (int i = 0; i < list.Count; i++)
                args[i] = list[i];
            return args;
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

        private List<string> CreateListCopy(string[] args)
        {
            List<string> list = new List<string>();
            foreach (string item in args)
                list.Add(null != item ? item.Trim() : String.Empty);
            return list;
        }

        private void RemoveEmptySpaceFromList(List<string> list)
        {
            bool flag = true;
            while (flag)
            {
                flag = list.Remove(String.Empty);
            }
        }

        private void RemovePreQuotesFromList(List<string> list)
        {
            for (int i = 0; i < list.Count; i++)
            {
                list[i] = null != list[i] ? list[i].Trim() : String.Empty;
                if (list[i].StartsWith("/", StringComparison.InvariantCultureIgnoreCase))
                    list[i] = list[i].Substring("/".Length);
                else if (list[i].StartsWith("-", StringComparison.InvariantCultureIgnoreCase))
                    list[i] = list[i].Substring("-".Length);
            }
        }

        #endregion
    }
}
