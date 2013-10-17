using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.ConsoleMonitor
{
    /// <summary>
    /// Custom start location for the main window
    /// </summary>
    internal enum WindowCustomStartLocation
    {
        /// <summary>
        /// top left corner
        /// </summary>
        TopLeft = 0,

        /// <summary>
        /// top right corner
        /// </summary>
        TopRight = 1,

        /// <summary>
        /// bottom left corner
        /// </summary>
        BottomLeft = 2,

        /// <summary>
        /// bottom right corner
        /// </summary>
        BottomRight = 3,

        /// <summary>
        /// center positon
        /// </summary>
        Center = 4,

        /// <summary>
        /// position from last exit
        /// </summary>
        LastPosition = 5,

        /// <summary>
        /// maximized state
        /// </summary>
        Maximized = 6
    }
}
