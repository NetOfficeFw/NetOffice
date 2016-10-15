using System;
using System.Reflection;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace NetOffice.Tools.WndUtils
{
    internal static class ExcelApplicationWindow
    { 
        #region Imports
        
        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref object ppvObject); 

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        private static uint _objectID = 0xFFFFFFF0;
        private static Guid _dispatchID = new Guid("00020400-0000-0000-C000-000000000046");
      
        #endregion

        #region Methods

        /// <summary>
        /// Try get the com application proxy from application window handle
        /// </summary>
        /// <param name="hwnd">excel application window handle</param>
        /// <returns>com proxy or null</returns>
        internal static object GetApplicationProxyFromHandle(IntPtr hwnd)
        {
            if (null == hwnd)
                throw new ArgumentNullException("hwnd");

            try
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
            catch (Exception exception)
            {
                DebugConsole.Default.WriteException(exception);
                throw;
            }
        }

        /// <summary>
        /// Returns a list with application proxies
        /// </summary>
        /// <param name="hwnds">main window handles</param>
        /// <returns>list of application proxies</returns>
        internal static Misc.DisposableObjectList GetApplicationProxiesFromHandle(IntPtr[] hwnds)
        {
            if (null == hwnds)
                throw new ArgumentNullException("hwnds");

            try
            {
                Misc.DisposableObjectList result = new Misc.DisposableObjectList();
                foreach (var item in hwnds)
                {
                    object app = GetApplicationProxyFromHandle(item);
                    if(null != app)
                        result.Add(app);
                }
                return result;
            }
            catch (Exception exception)
            {
                DebugConsole.Default.WriteException(exception);
                throw;
            }
        }

        #endregion
    }
}
