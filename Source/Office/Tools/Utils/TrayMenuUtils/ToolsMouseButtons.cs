using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Specifies constants that define which mouse button was pressed.
    /// </summary>
    [Flags]
    public enum ToolsMouseButtons
    {
        /// <summary>
        /// No mouse button was pressed.
        /// </summary>
        None = 0,

        /// <summary>
        /// The left mouse button was pressed.
        /// </summary>     
        Left = 1048576,

        /// <summary>
        /// The right mouse button was pressed.
        /// </summary>     
        Right = 2097152,

        /// <summary>
        /// The middle mouse button was pressed.
        /// </summary>
        Middle = 4194304,

        /// <summary>
        /// The first XButton was pressed.
        /// </summary>
        XButton1 = 8388608,

        /// <summary>
        /// The second XButton was pressed.
        /// </summary>
        XButton2 = 16777216
    }
}
