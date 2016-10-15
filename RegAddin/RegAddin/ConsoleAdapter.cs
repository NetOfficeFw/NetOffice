using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace RegAddin
{
    internal static class ConsoleAdapter
    {
        /// <summary>
        /// Allocates a new console for the calling process.
        /// </summary>
        /// <returns>If the function succeeds, the return value is nonzero.
        /// If the function fails, the return value is zero.
        /// To get extended error information, call Marshal.GetLastWin32Error.</returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern bool AllocConsole();

        /// <summary>
        /// Detaches the calling process from its console
        /// </summary>
        /// <returns>If the function succeeds, the return value is nonzero.
        /// If the function fails, the return value is zero.
        /// To get extended error information, call Marshal.GetLastWin32Error.</returns>
        [DllImport("kernel32", SetLastError = true)]
        private static extern bool FreeConsole();

        /// <summary>
        /// Attaches the calling process to the console of the specified process.
        /// </summary>
        /// <param name="dwProcessId">[in] Identifier of the process, usually will be ATTACH_PARENT_PROCESS</param>
        /// <returns>If the function succeeds, the return value is nonzero.
        /// If the function fails, the return value is zero.
        /// To get extended error information, call Marshal.GetLastWin32Error.</returns>
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AttachConsole(uint dwProcessId);

        /// <summary>Identifies the console of the parent of the current process as the console to be attached.</summary>
        private const uint ATTACH_PARENT_PROCESS = 0xFFFFFFFF;

        /// <summary>
        /// calling process is already attached to a console
        /// </summary>
        private const int ERROR_ACCESS_DENIED = 5;

        /// <summary>
        /// Allocate a console if application started from within windows GUI.
        /// Detects the presence of an existing console associated with the application and
        /// attaches itself to it if available.
        /// </summary>
        internal static bool TryAttach()
        {
            if (!AttachConsole(ATTACH_PARENT_PROCESS) && Marshal.GetLastWin32Error() == ERROR_ACCESS_DENIED)
            {
                if (!AllocConsole())
                {
                    return false;
                }
            }
            return true;
        }
    }
}
