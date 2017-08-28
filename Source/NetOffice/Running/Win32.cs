using System;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using COMTypes = System.Runtime.InteropServices.ComTypes;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Running
{
    /// <summary>
    /// Encapsulate some external Win32 operations to deal with windows desktop
    /// </summary>
    internal static class Win32
    {
        #region Imports
        
        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref object ppvObject);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern void GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("USER32.DLL")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        #endregion

        #region Fields

        private static uint _objectID = 0xFFFFFFF0;
        private static Guid _dispatchID = new Guid("00020400-0000-0000-C000-000000000046");
        private static int _builderLength = 128;

        #endregion

        #region Methods

        /// <summary>
        /// Retrieves the name of the class to which the specified window belongs. 
        /// </summary>
        /// <param name="hwnd">A handle to the window and, indirectly, the class to which the window belongs. </param>
        /// <returns>If the function succeeds, the return value is the number of characters copied to the buffer, not including the terminating null character.</returns>
        public static string GetClassName(IntPtr hwnd)
        {
            if (IntPtr.Zero == hwnd)
                return String.Empty;
            StringBuilder builder = new StringBuilder(_builderLength);
            GetClassName(hwnd, builder, _builderLength);
            return builder.ToString();
        }

        /// <summary>
        /// Copies the text of the specified window's title bar (if it has one) into a buffer. If the specified window is a control, the text of the control is copied. However, GetWindowText cannot retrieve the text of a control in another application.
        /// </summary>
        /// <param name="hwnd">A handle to the window or control containing the text. </param>
        /// <returns>The buffer that will receive the text. If the string is as long or longer than the buffer, the string is truncated and terminated with a null character.</returns>
        public static string GetWindowText(IntPtr hwnd)
        {
            if (IntPtr.Zero == hwnd)
                return String.Empty;
            StringBuilder builder = new StringBuilder(_builderLength);
            GetWindowText(hwnd, builder, _builderLength);
            return builder.ToString();
        }

        /// <summary>
        /// Retrieves the identifier of the thread that created the specified window and, optionally, the identifier of the process that created the window. 
        /// </summary>
        /// <param name="hwnd">A handle to the window. </param>
        /// <returns>The return value is the identifier of the thread that created the window.</returns>
        public static IntPtr GetWindowThreadProcessId(IntPtr hwnd)
        {
            if (IntPtr.Zero == hwnd)
                return IntPtr.Zero;
            uint procId = 0;
            GetWindowThreadProcessId(hwnd, out procId);
            if (procId != 0)
                return new IntPtr(procId);
            else
                return IntPtr.Zero;
        }

        /// <summary>
        /// Retrieves the address of the specified interface for the object associated with the specified window.
        /// </summary>
        /// <param name="hwnd">Specifies the handle of a window for which an object is to be retrieved. </param>
        /// <returns>Address of a pointer variable that receives the address of the specified interface.</returns>
        public static object AccessibleObjectFromWindow(IntPtr hwnd)
        {
            if (IntPtr.Zero == hwnd)
                return null;

            object accObject = new object();
            int result = AccessibleObjectFromWindow(hwnd, _objectID, ref _dispatchID, ref accObject);
            if (result == 0)
                return accObject;
            else
                return null;
        }

        #endregion
    }
}
