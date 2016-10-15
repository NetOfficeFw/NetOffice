using System;
using System.Drawing;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Text;

namespace NetOffice.Tools.WndUtils
{
    /// <summary>
    /// Enumerate all child windows for a window handle
    /// </summary>
    public class ChildWindowEnumerator
    {
        #region Imports

        private delegate bool EnumWindowProc(IntPtr hwnd, IntPtr lParam);
        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumChildWindows(IntPtr window, EnumWindowProc callback, IntPtr lParam);
        [DllImport("USER32.DLL")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        #endregion

        #region Fields

        private IntPtr _handle;
        private string _textFilter;
        private static object _lockInstance = new object();
        private static ChildWindowEnumerator _currentInstance;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="handle">target window handle</param>
        public ChildWindowEnumerator(IntPtr handle)
        {
            if (IntPtr.Zero == handle)
                throw new ArgumentOutOfRangeException("handle");
            _handle = handle;
            Result = new List<IntPtr>();
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="handle">target window handle</param>
        /// <param name="textFilter">optional text filter to compare window text</param>
        public ChildWindowEnumerator(IntPtr handle, string textFilter)
        {
            if (IntPtr.Zero == handle)
                throw new ArgumentOutOfRangeException("handle");
            _handle = handle;
            _textFilter = textFilter;
            Result = new List<IntPtr>();
        }

        #endregion

        #region Properties

        private List<IntPtr> Result { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// Enumerates all child windows. WARNING: The method returns null if operation timeout is reached.
        /// </summary>
        /// <param name="milliSecondsTimeout">a timeout for the operation. when a window is busy or non responding these method freeze. you can handle this with the operation timeout</param>
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

        private void EnumerateWindowsAsync(object mre)
        {
            try
            {
                Result = GetAllChildHandles();
                (mre as ManualResetEvent).Set();
            }
            catch (Exception exception)
            {
                DebugConsole.Default.WriteException(exception);
            }
        }

        private List<IntPtr> GetAllChildHandles()
        {
            List<IntPtr> childHandles = new List<IntPtr>();

            GCHandle gcChildhandlesList = GCHandle.Alloc(childHandles);
            IntPtr pointerChildHandlesList = GCHandle.ToIntPtr(gcChildhandlesList);

            try
            {
                EnumWindowProc childProc = new EnumWindowProc(EnumWindow);
                EnumChildWindows(_handle, childProc, pointerChildHandlesList);
            }
            finally
            {
                gcChildhandlesList.Free();
            }

            return childHandles;
        }

        private bool EnumWindow(IntPtr hWnd, IntPtr lParam)
        {
            GCHandle gcChildhandlesList = GCHandle.FromIntPtr(lParam);

            if (gcChildhandlesList == null || gcChildhandlesList.Target == null)
            {
                return false;
            }

            if (!String.IsNullOrEmpty(_textFilter))
            {
                StringBuilder builder = new StringBuilder();
                GetWindowText(hWnd, builder, 16);
                string windowText = builder.ToString();
                if (windowText.Equals(_textFilter, StringComparison.InvariantCultureIgnoreCase))
                {
                    List<IntPtr> childHandles = gcChildhandlesList.Target as List<IntPtr>;
                    childHandles.Add(hWnd);
                }
            }
            else
            {
                List<IntPtr> childHandles = gcChildhandlesList.Target as List<IntPtr>;
                childHandles.Add(hWnd);
            }

            return true;
        }

        #endregion
    }
}
