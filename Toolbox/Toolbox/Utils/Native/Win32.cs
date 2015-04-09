using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace NetOffice.DeveloperToolbox.Utils.Native
{
    internal static class Win32
    {
        public const int HWND_BROADCAST = 0xffff;
        public static readonly int WM_SHOWTOOLBOX = RegisterWindowMessage("WM_SHOWTOOLBOX");
        
        [DllImport("user32")]
        public static extern bool PostMessage(IntPtr hwnd, int msg, IntPtr wparam, IntPtr lparam);
        [DllImport("user32")]
        public static extern int RegisterWindowMessage(string message);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
    }
}
