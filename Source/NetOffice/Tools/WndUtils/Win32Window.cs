using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Tools.WndUtils
{
    /// <summary>
    /// IWin32Window Implementation
    /// </summary>
    public class Win32Window : System.Windows.Forms.IWin32Window
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="hwnd">target handle</param>
        public Win32Window(int hwnd)
        {
            Handle = new IntPtr(hwnd);
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="handle">target handle</param>
        public Win32Window(IntPtr handle)
        {
            Handle = handle;
        }

        /// <summary>
        /// Gets the handle to the window
        /// </summary>
        public IntPtr Handle { get; private set; }

        /// <summary>
        /// Try create an IWin32Window implementation by given argument
        /// </summary>
        /// <param name="value">target handle</param>
        /// <returns>IWin32Window</returns>
        public static System.Windows.Forms.IWin32Window Create(object value)
        {
            if (null == value)
                return null;

            System.Windows.Forms.IWin32Window wnd = value as System.Windows.Forms.IWin32Window;
            if (null != wnd)
                return new Win32Window(wnd.Handle);
            if (value is IntPtr)
                return new Win32Window((IntPtr)value);
            if (value is int)
                return new Win32Window((int)value);

            return null;
        }
    }
}
