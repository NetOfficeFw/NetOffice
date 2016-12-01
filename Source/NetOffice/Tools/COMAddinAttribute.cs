using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Common Addin Load Behavior Values
    /// </summary>
    public enum LoadBehavior
    {
        /// <summary>
        /// Do not load the addin
        /// </summary>
        DoNotLoad = 0,

        /// <summary>
        /// Load addin while startup
        /// </summary>
        LoadAtStartup = 3
    }

    /// <summary>
    /// Specify essential COMAddin informations
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Class, AllowMultiple = false)]
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
        /// <param name="loadBehavior">LoadBehavior for the Addin</param>
        public COMAddinAttribute(string name, LoadBehavior loadBehavior)
        {
            Name = name;
            Description = String.Empty;
            LoadBehavior = Convert.ToInt32(loadBehavior);
            CommandLineSafe = -1;
        }

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        /// <param name="name">Display Name for the Addin</param>
        /// <param name="description">Description for the Addin</param>
        /// <param name="loadBehavior">LoadBehavior for the Addin</param>
        public COMAddinAttribute(string name, string description, LoadBehavior loadBehavior)
        {
            Name = name;
            Description = description;
            LoadBehavior = Convert.ToInt32(loadBehavior);
            CommandLineSafe = -1;
        }

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
