using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.PowerPointApi.Tools.Utils
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
        private static extern int AccessibleObjectFromWindow(int hwnd, uint dwObjectID, byte[] riid, ref object ptr);

        private delegate bool EnumChildCallback(IntPtr hwnd, ref int lParam);

        #endregion

        #region Fields

        private static uint objid_NATIVEOM = 0xFFFFFFF0;
        private static Guid _dispatch = new Guid("{00020400-0000-0000-C000-000000000046}");
        private static Guid _unknown = new Guid("00000000-0000-0000-C000-000000000046");

        private CommonUtils _owner;
        private int _hwnd = 0;

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

        #region Properties

        /// <summary>
        /// Host application main window handle
        /// </summary>
        public int HWND
        {
            get
            {
                if (0 == _hwnd)
                    _hwnd = TryGetHostApplicationWindowHandle();
                return _hwnd;
            }
        }

        #endregion

        private bool EnumChildProc(IntPtr hwnd, ref int lParam)
        {
            StringBuilder windowClass = new StringBuilder(128);
            GetClassName(hwnd, windowClass, 128);
            if (windowClass.ToString() == "paneClassDC")
                lParam = (int)hwnd;
            return true;
        }

        #region Methods

        private int TryGetHostApplicationWindowHandle()
        {
            try
            {
                int result = 0;
                NetOffice.Tools.WndUtils.WindowEnumerator enumerator = new NetOffice.Tools.WndUtils.WindowEnumerator("PP", "FrameClass");
                IntPtr[] handles = enumerator.EnumerateWindows(2000);

                foreach (IntPtr item in handles)
                {
                    object proxyApplication = GetAccessibleObject(item);
                    if (null == proxyApplication)
                    {
                        bool equals = Equals(_owner.Owner.AppInstance.UnderlyingObject, proxyApplication);
                        if (equals)
                        {
                            result = (int)item;
                            Marshal.ReleaseComObject(proxyApplication);
                            break;
                        }
                        Marshal.ReleaseComObject(proxyApplication);
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
                    int hr = AccessibleObjectFromWindow(hWndChild, objid_NATIVEOM, _dispatch.ToByteArray(), ref ptr);
                    if (hr >= 0)
                        return NetOffice.Core.Default.Invoker.PropertyGet(ptr, "Application");
                }
            }

            return null;
        }

        private bool Equal(object applicationProxyA, object applicationProxyB)
        {
            IntPtr outValueA = IntPtr.Zero;
            IntPtr outValueB = IntPtr.Zero;
            IntPtr ptrA = IntPtr.Zero;
            IntPtr ptrB = IntPtr.Zero;
            try
            {
                ptrA = Marshal.GetIUnknownForObject(applicationProxyA);
                int hResultA = Marshal.QueryInterface(ptrA, ref _unknown, out outValueA);

                ptrB = Marshal.GetIUnknownForObject(applicationProxyB);
                int hResultB = Marshal.QueryInterface(ptrB, ref _unknown, out outValueB);

                return (hResultA == 0 && hResultB == 0 && ptrA == ptrB);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (IntPtr.Zero != ptrA)
                    Marshal.Release(ptrA);

                if (IntPtr.Zero != outValueA)
                    Marshal.Release(outValueA);

                if (IntPtr.Zero != ptrB)
                    Marshal.Release(ptrB);

                if (IntPtr.Zero != outValueB)
                    Marshal.Release(outValueB);
            }
        }

        #endregion
    }
}
