using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Text;

namespace ObjectFromWindow
{
    /// <summary>
    /// Enumerate Top Level Windows on Desktop
    /// </summary>
    public class WindowEnumerator
    {
        #region Imports
        
        protected delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        protected static extern int GetWindowText(IntPtr hWnd, StringBuilder strText, int maxCount);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        protected static extern int GetWindowTextLength(IntPtr hWnd);
        [DllImport("user32.dll")]
        protected static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
        [DllImport("user32.dll")]
        protected static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        protected static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

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
            Filter = filter;
            Result = new List<IntPtr>();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Optional class name filter or null
        /// </summary>
        public string Filter { get; private set; }

        private List<IntPtr> Result { get; set; }

        #endregion

        #region Methods
        
        /// <summary>
        /// Enumerates all top level windows on desktop. WARNING: The method returns null if operation timeout is reached.
        /// </summary>
        /// <param name="milliSecondsTimeout">a timeout for the operation. when a desktop is busy or non responding these method freeze. you can handle this with the operation timeout</param>
        /// <returns>Result Array or null</returns>
        public IntPtr[] EnumerateWindows(int milliSecondsTimeout)
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
                    return null;
                }
            }
            return Result.ToArray();
        }

        private void EnumerateWindowsAsync(object mre)
        {
            EnumWindows(new EnumWindowsProc(EnumTopLevelWindows), IntPtr.Zero);
            (mre as ManualResetEvent).Set();
        }

        private static bool EnumTopLevelWindows(IntPtr hWnd, IntPtr lParam)
        {
            int size = GetWindowTextLength(hWnd);
            if (size++ > 0)
            {
                StringBuilder sb = new StringBuilder(size);
                int nRet;
                StringBuilder sb2 = new StringBuilder(100);
                nRet = GetClassName(hWnd, sb2, sb2.Capacity);
                if(nRet != 0)
                {
                    string className = sb2.ToString();
                    if (null != _currentInstance.Filter)
                    {
                        if (_currentInstance.Filter.Equals(className, StringComparison.InvariantCultureIgnoreCase))
                            _currentInstance.Result.Add(hWnd);
                    }
                    else
                        _currentInstance.Result.Add(hWnd);
                }                
            }
            return true;
        } 

        #endregion
    }
}
