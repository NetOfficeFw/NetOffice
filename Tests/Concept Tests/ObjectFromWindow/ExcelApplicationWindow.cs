using System;
using System.Reflection;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace ObjectFromWindow
{
    public static class ExcelApplicationWindow
    {
        #region Imports
        
        [DllImport("oleacc.dll")]
        internal static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref object ppvObject); 

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        static uint _objectID = 0xFFFFFFF0;
        static Guid _dispatchID = new Guid("{00020400-0000-0000-C000-000000000046}");
      
        #endregion

        #region Methods

        /// <summary>
        /// Try get the com application proxy from application window handle
        /// </summary>
        /// <param name="hwnd">excel application window handle</param>
        /// <returns>com proxy or null</returns>
        public static object GetApplicationProxyFromHandle(IntPtr hwnd)
        {
            IntPtr hwnd2 = FindWindowEx(hwnd, IntPtr.Zero, "XLDESK", null);
            if (hwnd2 == (IntPtr)0)
                return null;
            IntPtr hwnd3 = FindWindowEx(hwnd2, IntPtr.Zero, "EXCEL7", null);
            if (hwnd3 == (IntPtr)0)
                return null;

            object accObject = new object();
            if (hwnd3 != (IntPtr)0)
            {
                AccessibleObjectFromWindow(hwnd3, _objectID, ref _dispatchID, ref accObject);
                if (accObject is MarshalByRefObject)
                { 
                    object targetProxy = accObject.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, accObject, new object[0]);
                    Marshal.ReleaseComObject(accObject);
                    return targetProxy;
                }
            }
            return null;
        }

        #endregion

    }
}
