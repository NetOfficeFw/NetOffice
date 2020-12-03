using System;

namespace NetOffice.Tools
{
    /// <summary>
    /// A value that specifies when the application attempts to load the add-in and the current state of the add-in (loaded or unloaded).
    /// By default add-ins use value <see cref="LoadAtStartup"/> (3), which specifies that the add-in is loaded at startup.
    /// </summary>
    /// <remarks>
    /// The LoadBehavior value contains a bitwise combination of values that specify the runtime behavior of the add-in.
    /// The lowest order bit (values 0 and 1) indicates whether the add-in is currently unloaded or loaded.
    /// Other bits indicate when the application attempts to load the add-in.
    /// Typically, the LoadBehavior entry is intended to be set to 0, 3, or 16 (in decimal) when the add-in is installed on end-user computer.
    /// </remarks>
    /// <see href="https://docs.microsoft.com/en-us/visualstudio/vsto/registry-entries-for-vsto-add-ins?view=vs-2019#LoadBehavior">LoadBehavior values</see>
    public enum LoadBehavior
    {
        /// <summary>
        /// The application never tries to load the add-in automatically.
        /// The user can try to manually load the add-in, or the add-in can be loaded programmatically.
        /// </summary>
        /// <remarks>
        /// If the add-in is successfully loaded, the LoadBehavior value remains 0.
        /// </remarks>
        DoNotLoad = 0,

        /// <summary>
        /// The application never tries to load the add-in automatically.
        /// The user can try to manually load the add-in, or the add-in can be loaded programmatically.
        /// </summary>
        /// <remarks>
        /// If the application successfully loads the add-in, the LoadBehavior value changes to 0, and remains at 0 after the application closes.
        /// </remarks>
        DoNotLoadManual = 1,

        /// <summary>
        /// The application does not try to load the add-in automatically.
        /// The user can try to manually load the add-in, or the add-in can be loaded programmatically.
        /// </summary>
        /// <remarks>
        /// If the application successfully loads the add-in, the LoadBehavior value changes to 3, and remains at 3 after the application closes.
        /// </remarks>
        LoadAtStartupManual = 2,

        /// <summary>
        /// The application tries to load the add-in when the application starts.
        /// This is the default value for most add-ins.
        /// </summary>
        /// <remarks>
        /// If the application successfully loads the add-in, the LoadBehavior value remains 3.
        /// If an error occurs when loading the add-in, the LoadBehavior value changes to 2, and remains at 2 after the application closes.
        /// </remarks>
        LoadAtStartup = 3,

        /// <summary>
        /// The application does not try to load the add-in automatically.
        /// The user can try to manually load the add-in, or the add-in can be loaded programmatically.
        /// </summary>
        /// <remarks>
        /// If the application successfully loads the add-in, the LoadBehavior value changes to 9.
        /// </remarks>
        LoadOnDemandManual = 8,

        /// <summary>
        /// The add-in will be loaded only when the application requires it,
        /// such as when a user clicks a UI element that uses functionality in the add-in
        /// (for example, a custom button in the Ribbon).
        /// </summary>
        /// <remarks>
        /// If the application successfully loads the add-in, the LoadBehavior value remains 9,
        /// but the status of the add-in in the COM Add-ins dialog box is updated to indicate that the add-in is currently loaded.
        /// If an error occurs when loading the add-in, the LoadBehavior value changes to 8.
        /// </remarks>
        LoadOnDemand = 9,

        /// <summary>
        /// Set this value if you want your add-in to be loaded on demand.
        /// The application loads the add-in when the user runs the application for the first time.
        /// The next time the user runs the application, the application loads any UI elements that are defined by the add-in,
        /// but the add-in is not loaded until the user clicks a UI element that is associated with the add-in.
        /// </summary>
        /// <remarks>
        /// When the application successfully loads the add-in for the first time, the LoadBehavior value remains 16 while the add-in is loaded.
        /// After the application closes, the LoadBehavior value changes to 9.
        /// </remarks>
        LoadOnce = 16
    }

    /// <summary>
    /// Specify essential COMAddin information
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