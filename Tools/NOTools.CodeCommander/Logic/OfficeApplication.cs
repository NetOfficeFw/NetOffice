using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.CodeCommander.Logic
{
    /// <summary>
    /// identifier for the current host application
    /// </summary>
    internal enum OfficeApplication
    {
        /// <summary>
        /// MS Excel
        /// </summary>
        Excel = 0,

        /// <summary>
        /// MS Word
        /// </summary>
        Word = 1,
        
        /// <summary>
        /// MS Outlook
        /// </summary>
        Outlook = 2,

        /// <summary>
        /// MS PowerPoint
        /// </summary>
        PowerPoint = 3,

        /// <summary>
        /// MS Access
        /// </summary>
        Access = 4,

        /// <summary>
        /// MS Project
        /// </summary>
        Project = 5
    }
}
