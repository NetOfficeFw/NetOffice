using System;
using System.Drawing;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Text;

namespace NetOffice.Tools.WndUtils
{
    /// <summary>
    /// Enumerate Top Level Windows on Desktop
    /// </summary>
    public class WindowEnumerator
    {
        #region Embedded Types

        /// <summary>
        /// Internal operatotion mode 
        /// </summary>
        public enum FilterMode
        {
            /// <summary>
            ///  Class name must match totaly 
            /// </summary>
            Full = 0,

            /// <summary>
            /// Class name must match in start
            /// </summary>
            Start = 2,

            /// <summary>
            /// Class name must match in end
            /// </summary>
            End = 3,

            /// <summary>
            /// Class name must match in start and end of name
            /// </summary>
            StartEnd = 1,
        }

        #endregion

        #region Imports

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top; 
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder strText, int maxCount);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextLength(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        #endregion

        #region Fields

        private static object _lockInstance = new object();
        private static WindowEnumerator _currentInstance;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="filter">optional class name filter or null</param>
        public WindowEnumerator(string filter)
        {
            Mode = FilterMode.Full;
            Filter = filter;
            StartsWithFilter = filter;
            EndsWithFilter = filter;
            Result = new List<IntPtr>();
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="filter">optional class name filter or null</param>
        /// <param name="mode">current filter mode</param>
        public WindowEnumerator(string filter, FilterMode mode)
        {
            Mode = mode;
            Filter = filter;
            StartsWithFilter = filter;
            EndsWithFilter = filter;
            Result = new List<IntPtr>();
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="startsWithfilter">starts with class name filter</param>
        /// <param name="endsWithFilter">ends with class name filter</param>
        public WindowEnumerator(string startsWithfilter, string endsWithFilter)
        {
            Mode = FilterMode.StartEnd;
            Filter = startsWithfilter;
            StartsWithFilter = startsWithfilter;
            EndsWithFilter = endsWithFilter;
            Result = new List<IntPtr>();
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="startsWithfilter">starts with class name filter</param>
        /// <param name="endsWithFilter">ends with class name filter</param>
        /// <param name="mode">current filter mode</param>
        public WindowEnumerator(string startsWithfilter, string endsWithFilter, FilterMode mode)
        {
            Mode = mode;
            Filter = startsWithfilter;
            StartsWithFilter = startsWithfilter;
            EndsWithFilter = endsWithFilter;
            Result = new List<IntPtr>();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Optional class name filter or null
        /// </summary>
        public string Filter { get; private set; }

        /// <summary>
        /// Class name begin
        /// </summary>
        public string StartsWithFilter { get; private set; }

        /// <summary>
        /// Class name end
        /// </summary>
        public string EndsWithFilter { get; private set; }

        /// <summary>
        /// Current Filter Mode
        /// </summary>
        public FilterMode Mode { get; private set; }

        private List<IntPtr> Result { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// Encapsulate WindowsAPI method GetWindowRect
        /// </summary>
        /// <param name="hwnd">target window handle</param>
        /// <returns>window coordinates dimensions to screen</returns>
        public static Rectangle GetWindowRect(IntPtr hwnd)
        {
            RECT rect = new RECT();
            GetWindowRect(hwnd, out rect);
            return new Rectangle(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);
        }

        /// <summary>
        /// Enumerates all top level windows on desktop. WARNING: The method returns null if operation timeout is reached.
        /// </summary>
        /// <param name="milliSecondsTimeout">a timeout for the operation. when a desktop windows is busy or non responding these method freeze. you can handle this with the operation timeout</param>
        /// <returns>result array or null</returns>
        public IntPtr[] EnumerateWindows(int milliSecondsTimeout)
        {
            if (milliSecondsTimeout < 0)
                throw new ArgumentOutOfRangeException("milliSecondsTimeout");

            try
            {
                lock (_lockInstance)
                {
                    Result.Clear();
                    _currentInstance = this;
                    Thread thread1 = new Thread(new ParameterizedThreadStart(EnumerateWindowsAsync));
                    WaitHandle[] waitHandles = new WaitHandle[1];
                    ManualResetEvent mre1 = new ManualResetEvent(false);
                    waitHandles[0] = mre1;
                    thread1.Start(mre1);
                    bool result = WaitHandle.WaitAll(waitHandles, milliSecondsTimeout);
                    if (!result)
                    {
                        thread1.Abort();
                        Result.Clear();
                        _currentInstance = null;
                        return null;
                    }
                    else
                    {
                        _currentInstance = null;
                    }
                }
                return Result.ToArray();
            }
            catch (Exception exception)
            {
                DebugConsole.Default.WriteException(exception);
                throw;
            }
        }

        /// <summary>
        /// Returns information a window is currently visible
        /// </summary>
        /// <param name="handle">target window handle</param>
        /// <returns>true if window is visible, otherwise false</returns>
        public bool IsVisible(IntPtr handle)
        {
            if (IntPtr.Zero == handle)
                throw new ArgumentNullException("handle");
            return IsWindowVisible(handle);
        }

        private void EnumerateWindowsAsync(object mre)
        {
            try
            {
                EnumWindows(new EnumWindowsProc(EnumTopLevelWindows), IntPtr.Zero);
                (mre as ManualResetEvent).Set();
            }
            catch (Exception exception)
            {
                DebugConsole.Default.WriteException(exception);
            }
        }

        private static bool EnumTopLevelWindows(IntPtr hWnd, IntPtr lParam)
        {
            try
            {
                int size = GetWindowTextLength(hWnd);
                if (size++ > 0)
                {
                    StringBuilder sb = new StringBuilder(size);
                    int nRet;
                    StringBuilder sb2 = new StringBuilder(100);
                    nRet = GetClassName(hWnd, sb2, sb2.Capacity);
                    if (nRet != 0)
                    {
                        string className = sb2.ToString();
                        if (FilterMatch(className, _currentInstance))
                            _currentInstance.Result.Add(hWnd);
                    }
                }
                return true;
            }
            catch (Exception exception)
            {
                DebugConsole.Default.WriteException(exception);
                return true;
            }
        }

        private static bool FilterMatch(string className, WindowEnumerator instance)
        {
            switch (instance.Mode)
            {
                case FilterMode.Full:
                {                       
                    if(false == String.IsNullOrEmpty(_currentInstance.Filter))
                    {
                        if (_currentInstance.Filter.Equals(className, StringComparison.InvariantCultureIgnoreCase))
                            return true;
                        else
                            return false;
                    }
                    else
                        return true;                    
                }
                case FilterMode.Start:
                {
                    string start = null != instance.StartsWithFilter ? instance.StartsWithFilter.ToLower() : "";                    
                    string target = className.ToLower();
                    return target.StartsWith(start);
                }
                case FilterMode.End:
                {
                    string end = null != instance.EndsWithFilter ? instance.EndsWithFilter.ToLower() : "";
                    string target = className.ToLower();
                    return target.EndsWith(end);
                }
                case FilterMode.StartEnd:
                {
                    string start = null != instance.StartsWithFilter ? instance.StartsWithFilter.ToLower() : "";
                    string end = null != instance.EndsWithFilter ? instance.EndsWithFilter.ToLower() : "";
                    string target = className.ToLower();    
                    return target.StartsWith(start) && target.EndsWith(end);
                }
                default:
                {
                    throw new IndexOutOfRangeException();
                }
            }           
        }

        #endregion
    }
}
