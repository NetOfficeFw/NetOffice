using System;
using System.Runtime.InteropServices;

namespace NetOffice.OutlookApi.Tools.Contribution.Security
{
    /// <summary>
    /// Simple RECT wrapper
    /// </summary>
    public class Rect
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="x">left</param>
        /// <param name="y">top</param>
        /// <param name="right">right bound</param>
        /// <param name="bottom">bottom bound</param>
        public Rect(int x, int y, int right, int bottom)
        {
            Left = x;
            Top = y;
            Right = right;
            Bottom = bottom;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="rect">inner api rect</param>
        internal Rect(User32.RECT rect)
        {
            Left = rect.Left;
            Top = rect.Top;
            Right = rect.Right;
            Bottom = rect.Bottom;
        }

        /// <summary>
        /// x position of upper-left corner
        /// </summary>
        public int Left;

        /// <summary>
        ///  y position of upper-left corner
        /// </summary>
        public int Top;

        /// <summary>
        ///  x position of lower-right corner
        /// </summary>
        public int Right;

        /// <summary>
        /// y position of lower-right corner
        /// </summary>
        public int Bottom;
    }
}
