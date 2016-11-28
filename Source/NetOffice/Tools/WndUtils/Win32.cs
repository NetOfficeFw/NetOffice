using System;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using COMTypes = System.Runtime.InteropServices.ComTypes;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Tools.WndUtils
{
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

        public static string GetClassName(IntPtr hwnd)
        {
            if (IntPtr.Zero == hwnd)
                return String.Empty;
            StringBuilder builder = new StringBuilder(_builderLength);
            GetClassName(hwnd, builder, _builderLength);
            return builder.ToString();
        }

        public static string GetWindowText(IntPtr hwnd)
        {
            if (IntPtr.Zero == hwnd)
                return String.Empty;
            StringBuilder builder = new StringBuilder(_builderLength);
            GetWindowText(hwnd, builder, _builderLength);
            return builder.ToString();
        }

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
