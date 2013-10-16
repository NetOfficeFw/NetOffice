using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Specify essential COMAddin informations
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Class)]
    public class COMAddinAttribute : System.Attribute
    {
        /// <summary>
        /// Display Name for the Addin
        /// </summary>
        public readonly string Name;

        /// <summary>
        /// Description for the Addin
        /// </summary>
        public readonly string Description;

        /// <summary>
        /// LoadBehavior for the Addin
        /// </summary>
        public readonly int LoadBehavior;

        /// <summary>
        /// Gives info the Addin is commandline safe
        /// </summary>
        public readonly int CommandLineSafe;

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        /// <param name="name">Display Name for the Addin</param>
        /// <param name="description">Description for the Addin</param>
        /// <param name="loadBehavior">LoadBehavior for the Addin</param>
        public COMAddinAttribute(string name, string description, int loadBehavior)
        {
            Name = name;
            Description = description;
            LoadBehavior = loadBehavior;
            CommandLineSafe = -1;
        }

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        /// <param name="name">Display Name for the Addin</param>
        /// <param name="description">Description for the Addin</param>
        /// <param name="loadBehavior">LoadBehavior for the Addin</param>
        /// <param name="commandLineSafe">Gives info the Addin is commandline safe</param>
        public COMAddinAttribute(string name, string description, int loadBehavior, int commandLineSafe)
        {
            Name = name;
            Description = description;
            LoadBehavior = loadBehavior;
            CommandLineSafe = commandLineSafe;
        }
    }
}
