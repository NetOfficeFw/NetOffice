using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace NetOffice.WordApi.Tools.Utils
{    
    /// <summary>
    /// Application related utils
    /// </summary>
    public class ApplicationUtils
    {
        #region Imports
        
        [DllImport("User32")]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("User32")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumChildCallback lpEnumFunc, ref int lParam);

        [DllImport("Oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(int hwnd, uint dwObjectID, byte[] riid, [MarshalAs(UnmanagedType.IDispatch)]ref object ptr);

        private delegate bool EnumChildCallback(IntPtr hwnd, ref int lParam);

        #endregion

        #region Fields

        private static uint objid_NATIVEOM = 0xFFFFFFF0;
        private static Guid _dispatch = new Guid("00020400-0000-0000-C000-000000000046");
        private static Guid _unknown = new Guid("00000000-0000-0000-C000-000000000046");

        private CommonUtils _owner;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">owner instance</param>
        protected internal ApplicationUtils(CommonUtils owner)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            _owner = owner;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Try to detect the main window handle for a document instance
        /// </summary>
        /// <param name="document">target document instance</param>
        /// <returns>main window handle or 0 if failed</returns>
        public int TryGetMainWindowHandle(WordApi.Document document)
        {
            if (null == document)
                throw new ArgumentNullException("document");
            int hwnd = TryGetHostApplicationWindowHandle(document);
            return hwnd;
        }

        private int TryGetHostApplicationWindowHandle(WordApi.Document document)
        {
            int result = TryGetHostApplicationWindowHandleFromDesktop(document);
            return result;
        }

        private int TryGetHostApplicationWindowHandleFromDesktop(WordApi.Document document)
        {
            try
            {
                int result = 0;
                NetOffice.Tools.WndUtils.WindowEnumerator enumerator = new NetOffice.Tools.WndUtils.WindowEnumerator("OpusApp");
                IntPtr[] handles = enumerator.EnumerateWindows(2000);
             
                foreach (IntPtr item in handles)
                {
                    object proxyDocument = GetAccessibleObject(item);
                    if (null != proxyDocument)
                    {
                        try
                        {
                            bool equals = Equal(document.UnderlyingObject, proxyDocument);
                            if (equals)
                                result = (int)item;
                            break;
                        }
                        catch
                        {
                            throw;
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(proxyDocument);
                        }
                    }
                }

                return result;
            }
            catch (Exception exception)
            {
                NetOffice.Core.Default.Console.WriteException(exception);
                return 0;
            }
        }

        private object GetAccessibleObject(IntPtr hwnd)
        {
            if (hwnd != IntPtr.Zero)
            {
                int hWndChild = 0;

                EnumChildCallback cb = new EnumChildCallback(EnumChildProc);
                EnumChildWindows(hwnd, cb, ref hWndChild);

                if (hWndChild != 0)
                {
                    object ptr = null;
                    object document = null;
                    int hr = AccessibleObjectFromWindow(hWndChild, objid_NATIVEOM, _dispatch.ToByteArray(), ref ptr);
                    if (hr >= 0)
                        document =  NetOffice.Core.Default.Invoker.PropertyGet(ptr, "Document");
                    Marshal.ReleaseComObject(ptr);
                    if (null != document)
                        return document;
                }
            }

            return null;
        }

        private bool EnumChildProc(IntPtr hwnd, ref int lParam)
        {
            StringBuilder windowClass = new StringBuilder(128);
            GetClassName(hwnd, windowClass, 128);
            if (windowClass.ToString() == "_WwG")
                lParam = (int)hwnd;
            return true;
        }

        private bool Equal(object applicationProxyA, object applicationProxyB)
        {
            try
            {
                COMObject a = new COMObject(applicationProxyA);
                COMObject b = new COMObject(applicationProxyB);
                return a.EqualsOnServer(b);
            }
            catch (Exception)
            {
                throw;
            }
        }

        #endregion
    }
}
