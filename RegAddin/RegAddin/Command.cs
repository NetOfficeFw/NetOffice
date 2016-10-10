using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin
{
    public enum AssemblyRequired
    {
        No = 0,
        Yes = 1,
        Conditional = 2
    }

    /// <summary>
    /// Represents a possible command in the application
    /// </summary>
    internal class Command
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="name">name of the command</param>
        /// <param name="argumentsSyntax">arguments description of the command</param>
        /// <param name="helpText">help text of the command</param>
        /// <param name="needAssemblyPath">the command need a path to an assembly</param>
        /// <param name="alias">additional alias names of the command</param>
        /// <param name="description">description/snytax validation of the command</param>
        internal Command(string name, string argumentsSyntax, string helpText, AssemblyRequired needAssemblyPath, IEnumerable<string> alias, CommandOptionDescriptions description)
        {
            if (null == name)
                throw new ArgumentNullException("name");
            if (null == helpText)
                throw new ArgumentNullException("helpText");
            Name = name;
            HelpText = helpText;
            NeedAssemblyPath = needAssemblyPath;
            Alias = null != alias ? alias : new string[0];
            ArgumentsSyntax = null != argumentsSyntax ? argumentsSyntax : "";
            Description = description;
        }
        
        /// <summary>
        /// Name of the command
        /// </summary>
        internal string Name { get; private set; }

        /// <summary>
        /// Arguments description of the command
        /// </summary>
        internal string ArgumentsSyntax { get; private set; }

        /// <summary>
        /// Additional alias names of the command
        /// </summary>
        internal IEnumerable<string> Alias { get; private set; }

        /// <summary>
        /// The command need a path to an assembly
        /// </summary>
        internal AssemblyRequired NeedAssemblyPath { get; private set; }

        /// <summary>
        /// Help text of the command
        /// </summary>
        internal string HelpText { get; private set; }

        /// <summary>
        /// Description/snytax validation of the command
        /// </summary>
        internal CommandOptionDescriptions Description { get; private set; }

        public override string ToString()
        {
            return Name;
        }
    }
}
