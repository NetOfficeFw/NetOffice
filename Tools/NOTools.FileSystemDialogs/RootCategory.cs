using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    /// <summary>
    /// The available root categories
    /// </summary>
    public enum RootCategory
    {
        /// <summary>
        ///  User Desktop
        /// </summary>
        Desktop = 0,

        /// <summary>
        /// Local Machine
        /// </summary>
        MyComputer = 1,

        /// <summary>
        /// User Files Folder
        /// </summary>
        MyDocuments = 2,

        /// <summary>
        /// SpecialFolders from System.IO
        /// </summary>
        SpecialFolders = 3,

        /// <summary>
        /// Custom defined folders
        /// </summary>
        TemplateFolders = 4,

        /// <summary>
        /// Not set
        /// </summary>
        Undefined
    }
}
