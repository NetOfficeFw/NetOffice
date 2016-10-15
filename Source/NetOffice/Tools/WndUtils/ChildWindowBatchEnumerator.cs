using System;
using System.Drawing;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Text;

namespace NetOffice.Tools.WndUtils
{
    /// <summary>
    /// Try to find specific child windows in a deep level
    /// </summary>
    public class ChildWindowBatchEnumerator
    {
        #region Imports

        private delegate bool EnumWindowProc(IntPtr hwnd, IntPtr lParam);
        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumChildWindows(IntPtr window, EnumWindowProc callback, IntPtr lParam);
        [DllImport("USER32.DLL")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        #endregion

        #region Nested

        /// <summary>
        /// Child window search criteria
        /// </summary>
        public class SearchCriteria
        {
            /// <summary>
            /// Creates an instance of the class
            /// </summary>
            public SearchCriteria()
            {

            }

            /// <summary>
            /// Creates an instance of the class
            /// </summary>
            /// <param name="className"></param>
            public SearchCriteria(string className)
            {
                ClassName = className;
            }

            /// <summary>
            /// Target window class name
            /// </summary>
            public string ClassName { get; private set; }

            /// <summary>
            /// Returns a System.String that represents the instance
            /// </summary>
            /// <returns>System.String</returns>
            public override string ToString()
            {
                return String.IsNullOrWhiteSpace(ClassName) ? base.ToString() : ClassName;
            }
        }

        #endregion

        #region Fields

        private IntPtr _rootHandle;
        private static object _lockInstance = new object();
        private static ChildWindowBatchEnumerator _currentInstance;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="handle">target window handle</param>
        public ChildWindowBatchEnumerator(IntPtr handle)
        {
            if (IntPtr.Zero == handle)
                throw new ArgumentOutOfRangeException("handle");
            _rootHandle = handle;
            Result = new List<IntPtr>();
            SearchOrder = new List<SearchCriteria>();
        }

        #endregion

        #region Properties
     
        /// <summary>
        /// Search order as top down batch
        /// </summary>
        public List<SearchCriteria> SearchOrder { get; private set; }

        private SearchCriteria Current { get; set; }

        private List<IntPtr> Result { get; set; }

        private List<IntPtr> LastStepResult { get; set; }

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

            if (SearchOrder.Count == 0)
                return new IntPtr[0];

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
            if (null == LastStepResult)
                LastStepResult = new List<IntPtr>();
            else
                LastStepResult.Clear();
            LastStepResult.Add(_rootHandle);

            foreach (SearchCriteria search in SearchOrder)
            {
                Current = search;

                List<IntPtr> childHandles = new List<IntPtr>();
                GCHandle gcChildhandlesList = GCHandle.Alloc(childHandles);
                IntPtr pointerChildHandlesList = GCHandle.ToIntPtr(gcChildhandlesList);

                foreach (IntPtr handle in LastStepResult)
                {
                    EnumWindowProc childProc = new EnumWindowProc(EnumWindow);
                    EnumChildWindows(handle, childProc, pointerChildHandlesList);
                }

                if (childHandles.Count == 0)
                    break;

                LastStepResult.Clear();
                foreach (IntPtr item in childHandles)
                    LastStepResult.Add(item);
                childHandles.Clear();
                gcChildhandlesList.Free();
            }

            Result.Clear();
            foreach (IntPtr item in LastStepResult)
                Result.Add(item);
            
            return Result;
        }

        private bool EnumWindow(IntPtr hWnd, IntPtr lParam)
        {
            GCHandle gcChildhandlesList = GCHandle.FromIntPtr(lParam);

            if (gcChildhandlesList == null || gcChildhandlesList.Target == null)
            {
                return false;
            }
             
            StringBuilder nameBuilder = new StringBuilder(100);
            int result = GetClassName(hWnd, nameBuilder, nameBuilder.Capacity);
            if (result != 0)
            {
                string className = nameBuilder.ToString();
                if(Current.ClassName.Equals(className, StringComparison.InvariantCultureIgnoreCase))
                {
                    List<IntPtr> childHandles = gcChildhandlesList.Target as List<IntPtr>;
                    childHandles.Add(hWnd);
                }
            }

            return true;
        }

        private SearchCriteria GetNextSearchCriteria()
        {
            if (SearchOrder.Count == 0)
                return null;
            if (null == Current)
                return SearchOrder[0];

            int currentIndex = SearchOrder.IndexOf(Current);
            if (currentIndex < SearchOrder.Count)
                return SearchOrder[currentIndex + 1];
            else
                return null;
        }

        #endregion
    }
}
