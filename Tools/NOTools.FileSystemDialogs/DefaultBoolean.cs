using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    /// <summary>
    /// Represents a defaultable boolean
    /// </summary>
    public enum DefaultBoolean
    {
        /// <summary>
        /// always false
        /// </summary>
        False = 0,

        /// <summary>
        /// always true
        /// </summary>
        True = 1,

        /// <summary>
        /// use default settings
        /// </summary>
        Default = 2
    }
}
