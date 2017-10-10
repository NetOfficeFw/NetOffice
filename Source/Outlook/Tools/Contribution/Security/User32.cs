using System.Collections.Generic;
using System.Runtime.InteropServices;
using System;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace NetOffice.OutlookApi.Tools.Contribution.Security
{
    /// <summary>
    /// User32.dll imports (taken and modified from pinvoke.net)
    /// </summary>
    internal class User32
    {
        #region Constants

        private static uint MOUSEEVENTF_ABSOLUTE = 0x8000;
        private static uint MOUSEEVENTF_MOVE = 0x0001;
        private static uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private static uint MOUSEEVENTF_LEFTUP = 0x0004;

        #endregion

        #region Structs

        [StructLayout(LayoutKind.Sequential)]
        internal struct RECT
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct POINT
        {
            public int X;
            public int Y;

            public POINT(int x, int y)
            {
                this.X = x;
                this.Y = y;
            }

            public static implicit operator System.Drawing.Point(POINT p)
            {
                return new System.Drawing.Point(p.X, p.Y);
            }

            public static implicit operator POINT(System.Drawing.Point p)
            {
                return new POINT(p.X, p.Y);
            }
        }

        #endregion

        #region Delegates

        /// <summary>
        /// Window enumerator 
        /// </summary>
        /// <param name="hWnd">entry handle</param>
        /// <param name="lParam">optionals</param>
        /// <returns>true if child wnd found</returns>
        internal delegate bool EnumDelegate(IntPtr hWnd, int lParam);

        /// <summary>
        /// Callback enumerator
        /// </summary>
        /// <param name="hWnd">entry handle</param>
        /// <param name="parameter">optionals</param>
        /// <returns>true if callback found</returns>
        internal delegate bool EnumWindowProc(IntPtr hWnd, IntPtr parameter);

        #endregion

        #region Externals

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        internal static extern bool GetCursorPos(ref Point lpPoint);

        [DllImport("user32.dll")]
        internal static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        internal static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool EnumChildWindows(IntPtr window, EnumWindowProc callback, IntPtr i);

        [DllImport("user32.dll", EntryPoint = "GetClassName", ExactSpelling = false, CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", EntryPoint = "GetWindowText", ExactSpelling = false, CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern int GetWindowText(IntPtr hWnd, StringBuilder lpWindowText, int nMaxCount);

        [DllImport("user32.dll", EntryPoint = "EnumDesktopWindows", ExactSpelling = false, CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern bool EnumDesktopWindows(IntPtr hDesktop, EnumDelegate lpEnumCallbackFunction, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsWindowEnabled(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool EnableWindow(IntPtr hWnd, bool bEnable);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        internal static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern int PostMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

        private const int BM_CLICK = 0x00F5;

        #endregion

        #region Methods

        private static bool EnumWindow(IntPtr handle, IntPtr pointer)
        {
            GCHandle gch = GCHandle.FromIntPtr(pointer);
            List<IntPtr> list = gch.Target as List<IntPtr>;
            if (list == null)
                throw new InvalidCastException("GCHandle Target could not be cast as List<IntPtr>");

            list.Add(handle);

            return true;
        }

        /// <summary>
        /// text value of window
        /// </summary>
        /// <param name="hWnd"></param>
        /// <returns></returns>
        public static string GetWindowText(IntPtr hWnd)
        {
            StringBuilder strbText = new StringBuilder(255);
            User32.GetWindowText(hWnd, strbText, strbText.Capacity + 1);
            return strbText.ToString();
        }

        /// <summary>
        /// Returns a list of child windows
        /// </summary>
        /// <param name="parent">Parent of the windows to return</param>
        /// <returns>List of child windows</returns>
        public static List<IntPtr> GetChildWindows(IntPtr parent)
        {
            List<IntPtr> result = new List<IntPtr>();
            GCHandle listHandle = GCHandle.Alloc(result);
            try
            {
                EnumWindowProc childProc = new EnumWindowProc(EnumWindow);
                EnumChildWindows(parent, childProc, GCHandle.ToIntPtr(listHandle));
            }
            finally
            {
                if (listHandle.IsAllocated)
                    listHandle.Free();
            }
            return result;
        }

        /// <summary>
        /// class name of window
        /// </summary>
        /// <param name="hWnd">target window handle</param>
        /// <returns>window class name</returns>
        public static string GetClassName(IntPtr hWnd)
        {
            StringBuilder strbClassName = new StringBuilder(255);
            User32.GetClassName(hWnd, strbClassName, strbClassName.Capacity + 1);
            return strbClassName.ToString();
        }

        /// <summary>
        /// Move the mouse to x:y,  do left mouse click and move mouse back to origin position
        /// </summary>
        /// <param name="x">location on x</param>
        /// <param name="y">location on y</param>
        public static void DoMouseMoveClick(int x, int y)
        {
            Point mousePoint = new Point();
            GetCursorPos(ref mousePoint);
            x = (int)(65535.0f / Screen.PrimaryScreen.Bounds.Width * x);
            y = (int)(65535.0f / Screen.PrimaryScreen.Bounds.Height * y);
            User32.mouse_event(User32.MOUSEEVENTF_ABSOLUTE + User32.MOUSEEVENTF_MOVE, (uint)x, (uint)y, 0, 0);
            User32.mouse_event(User32.MOUSEEVENTF_ABSOLUTE | User32.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            User32.mouse_event(User32.MOUSEEVENTF_ABSOLUTE | User32.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            SetCursorPos(mousePoint.X, mousePoint.Y);
        }

        /// <summary>
        /// Send BM_CLICK via SendMessage
        /// </summary>
        /// <param name="handle">target handle</param>
        public static void DoSendMouseClick(IntPtr handle)
        {
            SendMessage(handle, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
        }

        public static void DoPostMouseClick(IntPtr handle)
        {
            PostMessage(handle, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
        }

        #endregion
    }
}
