using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace NetOffice.DeveloperToolbox.Utils.Native
{
    /// <summary>
    /// Windows Message API utils
    /// </summary>
    internal static class Win32
    {
        /// <summary>
        /// Signals a broadcast to the environment
        /// </summary>
        public const int HWND_BROADCAST = 0xffff;

        /// <summary>
        /// Registered custom WM Message
        /// </summary>
        public static readonly int WM_SHOWTOOLBOX = RegisterWindowMessage("WM_SHOWTOOLBOX");
        
        /// <summary>
        /// Well known Win32 PostMessage method to post a message to a window directly
        /// </summary>
        /// <param name="hwnd">handle of the target window</param>
        /// <param name="msg">message kind like WM_BORING</param>
        /// <param name="wparam">first argument depending on msg</param>
        /// <param name="lparam">second argument depending on msg</param>
        /// <returns>true if send sucsessfuly, otherwise false</returns>
        [DllImport("user32")]
        public static extern bool PostMessage(IntPtr hwnd, int msg, IntPtr wparam, IntPtr lparam);

        /// <summary>
        /// Brings a native window in from
        /// </summary>
        /// <param name="hWnd">handle of the target window</param>
        /// <returns>true if send sucsessfuly(may this means not the window is in front), otherwise false</returns>
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        /// <summary>
        /// Register a more readable message code
        /// </summary>
        /// <param name="message">unique messag to register</param>
        /// <returns>true if its registered, otherwise false</returns>
        [DllImport("user32")]
        public static extern int RegisterWindowMessage(string message);
    }
}
