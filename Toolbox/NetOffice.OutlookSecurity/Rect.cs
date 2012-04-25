using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OutlookSecurity
{
    public struct Rect
    {
        public Rect(int x, int y, int right, int bottom)
        {
            Left = x;
            Top = y;
            Right = right;
            Bottom = bottom;
        }

        internal Rect(User32.RECT rect)
        {
            Left = rect.Left;
            Top = rect.Top;
            Right = rect.Right;
            Bottom = rect.Bottom;
        }

        public int Left;        // x position of upper-left corner
        public int Top;         // y position of upper-left corner
        public int Right;       // x position of lower-right corner
        public int Bottom;      // y position of lower-right corner
    }
}
