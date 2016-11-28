using System;
using System.Drawing;
using System.Runtime;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Provides data for mouse events.
    /// </summary>
    public class ToolsMouseEventArgs : EventArgs
    {
        private Point? _location;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="button">One of the mouse buttons values that indicate which mouse button was pressed.</param>
        /// <param name="clicks">The number of times a mouse button was pressed.</param>
        /// <param name="x">The x-coordinate of a mouse click, in pixels.</param>
        /// <param name="y">The y-coordinate of a mouse click, in pixels.</param>
        /// <param name="delta">A signed count of the number of detents the wheel has rotated.</param>                   
        [TargetedPatchingOptOut("Performance critical to inline this type of method across NGen image boundaries")]
        public ToolsMouseEventArgs(ToolsMouseButtons button, int clicks, int x, int y, int delta)
        {
            Button = button;
            Clicks = clicks;
            X = x;
            Y = y;
            Delta = delta;
        }

        /// <summary>
        /// >One of the mouse buttons values that indicate which mouse button was pressed.
        /// </summary>
        public ToolsMouseButtons Button { get; private set; }

        /// <summary>
        /// The number of times a mouse button was pressed.
        /// </summary>
        public int Clicks { get; private set; }

        /// <summary>
        /// A signed count of the number of detents the wheel has rotated.
        /// </summary>
        public int Delta { get; private set; }

        /// <summary>
        ///  Gets the location of the mouse during the generating mouse event.
        /// </summary>
        public Point Location
        {
            get
            {
                if (null == _location)
                    _location = new Point(X, Y);
                return _location.Value;
            }
        }

        /// <summary>
        /// The x-coordinate of a mouse click, in pixels.
        /// </summary>
        public int X { get; private set; }

        /// <summary>
        /// The y-coordinate of a mouse click, in pixels
        /// </summary>
        public int Y { get; private set; }
    }
}
