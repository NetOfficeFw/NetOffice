using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice;

namespace NOTools.CodeCommander.Logic
{
    /// <summary>
    /// Runtime commands must implement this interface
    /// </summary>
    public abstract class DynamicCommand
    {
        public DynamicCommand()
        {
            ShowDialogOnError = true;
        }

        /// <summary>
        /// Office Host Application
        /// </summary>
        public COMObject HostApplication { get; private set; }

        /// <summary>
        /// Show error dialog on error(exception) true by default
        /// </summary>
        public bool ShowDialogOnError { get; protected set; }

        /// <summary>
        /// Connect the command with the office host application
        /// </summary>
        /// <param name="hostApplication"></param>
        public void Initialize(NetOffice.COMObject hostApplication)
        { 
            HostApplication= hostApplication;
        }
        
        /// <summary>
        /// Execute the command
        /// </summary>
        public abstract void Execute();
    }
}
