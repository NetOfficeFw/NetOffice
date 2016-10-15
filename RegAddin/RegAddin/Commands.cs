using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin
{
    /// <summary>
    /// Set of possible options for a command. Optional options must be set at the end only like C# argument rules
    /// </summary>
    internal class CommandOptionDescriptions : IEnumerable<CommandOptionDescription>
    {
        private List<CommandOptionDescription> _items = new List<CommandOptionDescription>();

        public int Count
        {
            get
            {
                return _items.Count;
            }
        }

        internal void Add(CommandOptionDescription description)
        {
            if (null == description)
                throw new ArgumentException("description");

            if (false == description.IsOptional && false == IsValidOptionalOrNonOptional(description))
                throw new InvalidOperationException("Unable to add optional description for option/non-optional rules");

            _items.Add(description);
        }

        private bool IsValidOptionalOrNonOptional(CommandOptionDescription description)
        {
            if(description.IsOptional)
            {
                CommandOptionDescription lastCommand = TryGetLastCommand();
                CommandOptionDescription penultimateCommand = TryGetPenultimateCommand();
                if (null != lastCommand && null != penultimateCommand)
                {
                    if (lastCommand.IsOptional && penultimateCommand.IsOptional)
                        return false;
                }
                else
                {
                    if (null == penultimateCommand)
                        return false;
                }

                // optionals dürfen immer angefügt weden wenn der letze auch ein optional ist
                // oder dies der erste optional ist
            }
            else
            {
                CommandOptionDescription lastCommand = TryGetLastCommand();
                if (null != lastCommand && lastCommand.IsOptional)
                    return false;
            }

            return true;
        }

        private CommandOptionDescription TryGetPenultimateCommand()
        {
            if (_items.Count > 1)
                return _items[_items.Count - 2];
            else
                return null;
        }

        private CommandOptionDescription TryGetLastCommand()
        {
            if (_items.Count > 0)
                return _items[_items.Count - 1];
            else
                return null;
        }

        public IEnumerator<CommandOptionDescription> GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _items.GetEnumerator();
        }
    }

    /// <summary>
    /// Detailed description for a single command option
    /// </summary>
    internal class CommandOptionDescription
    {
        internal CommandOptionDescription(string name, bool isOptional, string defaultValueIfOptional)
        {
            Name = name;
            IsOptional = isOptional;
            DefaultValueIfOptional = defaultValueIfOptional;
        }

        internal CommandOptionDescription(string name, bool isOptional, string defaultValueIfOptional, bool isEnumOption, IEnumerable<string> possibleEnumValues)
        {
            Name = name;
            IsOptional = isOptional;
            DefaultValueIfOptional = defaultValueIfOptional;
            IsEnumOption = isEnumOption;
            PossibleEnumValues = possibleEnumValues;
        }

        /// <summary>
        /// The internal name of the description
        /// </summary>
        internal string Name { get; private set; }

        /// <summary>
        /// The value is optional(must not set) and a default value is used when its not defined
        /// </summary>
        internal bool IsOptional { get; private set; }

        /// <summary>
        /// The used default value if the option is optional
        /// </summary>
        internal string DefaultValueIfOptional { get; private set; }

        /// <summary>
        /// The option has a fixed set of possible values
        /// </summary>
        internal bool IsEnumOption { get; private set; }

        /// <summary>
        /// Possible fixed set of enum values when the command option is an enum option
        /// </summary>
        internal IEnumerable<string> PossibleEnumValues { get; private set; }

        internal bool IsValueValid(string expression)
        {
            if (IsEnumOption)
                return PossibleEnumValues.Any(e => e.Equals(expression, StringComparison.InvariantCultureIgnoreCase));
            else
                return true;
        }

        public override string ToString()
        {
            return Name;
        }
    }

    internal class Commands : List<Command>
    {
        internal Commands()
        {
            Add(new Command("diag", "", "Shows diagnostic message box", AssemblyRequired.No, null, null));

            CommandOptionDescriptions regOptions = new CommandOptionDescriptions();
            regOptions.Add(new CommandOptionDescription("RegisterArea", true, "system", true, new string[] { "system", "user"}));
            regOptions.Add(new CommandOptionDescription("RegisterCall", true, "on", true, new string[] { "off", "on" }));
            regOptions.Add(new CommandOptionDescription("CreateOfficeKeys", true, "off", true, new string[] { "off", "on" }));
            Add(new Command("reg", ":system|user:off|on:off|on", "Register Addin", AssemblyRequired.Yes, null, regOptions));
             
            CommandOptionDescriptions unregOptions = new CommandOptionDescriptions();
            unregOptions.Add(new CommandOptionDescription("UnRegisterArea", true, "auto", true, new string[] { "auto", "system", "user" }));
            unregOptions.Add(new CommandOptionDescription("UnRegisterCall", true, "on", true, new string[] { "off", "on" }));
            unregOptions.Add(new CommandOptionDescription("DeleteOfficeKeys", true, "off", true, new string[] { "off", "on" }));
            unregOptions.Add(new CommandOptionDescription("SuspendMissingAssemblyError", true, "true", true, new string[] { "true", "false" }));
            Add(new Command("unreg", ":system|user:off|on:off|on:true|false", "Unregister Addin", AssemblyRequired.Conditional, null, unregOptions));

            CommandOptionDescriptions regFileOptions = new CommandOptionDescriptions();
            regFileOptions.Add(new CommandOptionDescription("RegFileArea", false, "system", true, new string[] { "system", "user" }));
            regFileOptions.Add(new CommandOptionDescription("RegFileCall", false, "off", true, new string[] { "off", "on" }));
            regFileOptions.Add(new CommandOptionDescription("CreateOfficeKeys", false, "off", true, new string[] { "off", "on" }));
            regFileOptions.Add(new CommandOptionDescription("RegFilePath", false, null, false, null));
            Add(new Command("regfile", ":system|user:$path", "Generate a register(.reg)file", AssemblyRequired.Yes, null, regFileOptions));

            Add(new Command("help", ":console|window", "Show help content in console or messagebox", AssemblyRequired.No, new string[] { "?" }, null));
            Add(new Command("codebase", "", "Add or remove codebase entry in reg/unreg", AssemblyRequired.Conditional, null, null));

            CommandOptionDescriptions alertOptions = new CommandOptionDescriptions();
            alertOptions.Add(new CommandOptionDescription("AlertMode", true, "error", true, new string[] { "on", "off", "error" }));
            Add(new Command("alert", ":on|off|error", "Shows a message box at the end", AssemblyRequired.No, null, alertOptions));

            CommandOptionDescriptions signOptions = new CommandOptionDescriptions();
            signOptions.Add(new CommandOptionDescription("SignCheck", true, "off", true, new string[] { "warn", "error", "off" }));            
            Add(new Command("sign", ":warn|error|off", "Throws an error or warning when the given assembly is not signed", AssemblyRequired.No, null, signOptions));

            CommandOptionDescriptions metricsOptions = new CommandOptionDescriptions();
            metricsOptions.Add(new CommandOptionDescription("MetricsOptions", true, "none", true, new string[] { "none", "con", "win" }));
            Add(new Command("metrics", ":none|con|win", "Analyze assembly", AssemblyRequired.Conditional, null, metricsOptions));
             
            CommandOptionDescriptions anyOptions = new CommandOptionDescriptions();
            anyOptions.Add(new CommandOptionDescription("AnyTag", false, null, false, null));
            Add(new Command("any", ":$any", "Indentity call for troubleshooting", AssemblyRequired.No, null, anyOptions));
        }

        internal Command this[string name]
        {
            get
            {
                return this.First(e => e.Name == name);
            }
        }

        internal CommandsSyntax Syntax
        {
            get
            {
                CommandsSyntax syntax = new CommandsSyntax();

                CommandSyntax reg = new CommandSyntax(this["reg"]);
                reg.Items.Add(new CommandSyntax(this["codebase"]));               
                reg.Items.Add(new CommandSyntax(this["alert"]));
                reg.Items.Add(new CommandSyntax(this["sign"]));
                reg.Items.Add(new CommandSyntax(this["any"]));
                reg.Items.Add(new CommandSyntax(this["metrics"]));
                syntax.Add(reg);

                CommandSyntax unreg = new CommandSyntax(this["unreg"]);
                unreg.Items.Add(new CommandSyntax(this["codebase"]));
                unreg.Items.Add(new CommandSyntax(this["alert"]));
                unreg.Items.Add(new CommandSyntax(this["sign"]));
                unreg.Items.Add(new CommandSyntax(this["any"]));
                syntax.Add(unreg);

                CommandSyntax regFile = new CommandSyntax(this["regfile"]);
                regFile.Items.Add(new CommandSyntax(this["codebase"]));
                regFile.Items.Add(new CommandSyntax(this["alert"]));
                regFile.Items.Add(new CommandSyntax(this["sign"]));
                regFile.Items.Add(new CommandSyntax(this["any"]));
                syntax.Add(regFile);

                CommandSyntax help = new CommandSyntax(this["help"]);
                help.Items.Add(new CommandSyntax(this["any"]));
                syntax.Add(help);

                CommandSyntax diag = new CommandSyntax(this["diag"]);
                diag.Items.Add(new CommandSyntax(this["any"]));
                syntax.Add(diag);

                return syntax;
            }
        }     
    }
}
