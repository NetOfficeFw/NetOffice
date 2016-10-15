using System;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace NetOffice.Tools.WndUtils
{
    /// <summary>
    /// Find out another process has admin elevations
    /// </summary>
    public static class ProcessElevation
    {
        #region Fields

        private static uint STANDARD_RIGHTS_READ = 0x00020000;
        private static uint TOKEN_QUERY = 0x0008;
        private static uint TOKEN_READ = (STANDARD_RIGHTS_READ | TOKEN_QUERY);

        #endregion

        #region Imports
 
        [DllImport("advapi32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool OpenProcessToken(IntPtr ProcessHandle, UInt32 DesiredAccess, out IntPtr TokenHandle);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool CloseHandle(IntPtr hObject);

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool GetTokenInformation(IntPtr TokenHandle, TOKEN_INFORMATION_CLASS TokenInformationClass, IntPtr TokenInformation, uint TokenInformationLength, out uint ReturnLength);

        private enum TOKEN_INFORMATION_CLASS
        {
            TokenUser = 1,
            TokenGroups,
            TokenPrivileges,
            TokenOwner,
            TokenPrimaryGroup,
            TokenDefaultDacl,
            TokenSource,
            TokenType,
            TokenImpersonationLevel,
            TokenStatistics,
            TokenRestrictedSids,
            TokenSessionId,
            TokenGroupsAndPrivileges,
            TokenSessionReference,
            TokenSandBoxInert,
            TokenAuditPolicy,
            TokenOrigin,
            TokenElevationType,
            TokenLinkedToken,
            TokenElevation,
            TokenHasRestrictions,
            TokenAccessInformation,
            TokenVirtualizationAllowed,
            TokenVirtualizationEnabled,
            TokenIntegrityLevel,
            TokenUIAccess,
            TokenMandatoryPolicy,
            TokenLogonSid,
            MaxTokenInfoClass
        }

        private enum TOKEN_ELEVATION_TYPE
        {
            TokenElevationTypeDefault = 1,
            TokenElevationTypeFull,
            TokenElevationTypeLimited
        }

        #endregion

        #region Methods
        
        /// <summary>
        /// Converts nulllable bool to Process elevation
        /// </summary>
        /// <param name="value">target value to convert</param>
        /// <returns>process elevation</returns>
        public static ProxyInformation.ProcessElevation ConvertToProcessElevation(bool? value)
        {
            if (true == value)
                return ProxyInformation.ProcessElevation.AdministratorRole;
            else if (false == value)
                return ProxyInformation.ProcessElevation.BelowAdministratorRole;
            else
                return ProxyInformation.ProcessElevation.Unknown;
        }

        /// <summary>
        /// Returns information about process elevation
        /// </summary>
        /// <param name="processHandle">target process id</param>
        /// <returns>true if elevated, null if unknown</returns>
        public static bool? IsProcessElevated(IntPtr processHandle)
        {
            if (processHandle == IntPtr.Zero)
                return null;

            IntPtr tokenHandle = IntPtr.Zero;
            if (!OpenProcessToken(Process.GetCurrentProcess().Handle, TOKEN_READ, out tokenHandle))
            {
                return null;
            }

            try
            {
                TOKEN_ELEVATION_TYPE elevationResult = TOKEN_ELEVATION_TYPE.TokenElevationTypeDefault;

                int elevationResultSize = Marshal.SizeOf((int)elevationResult);
                uint returnedSize = 0;

                IntPtr elevationTypePtr = Marshal.AllocHGlobal(elevationResultSize);
                try
                {
                    bool success = GetTokenInformation(tokenHandle, TOKEN_INFORMATION_CLASS.TokenElevationType,
                                                       elevationTypePtr, (uint)elevationResultSize,
                                                       out returnedSize);
                    if (success)
                    {
                        elevationResult = (TOKEN_ELEVATION_TYPE)Marshal.ReadInt32(elevationTypePtr);
                        bool isProcessAdmin = elevationResult == TOKEN_ELEVATION_TYPE.TokenElevationTypeFull;
                        return isProcessAdmin;
                    }
                    else
                    {
                        return null;
                    }
                }
                finally
                {
                    if (elevationTypePtr != IntPtr.Zero)
                        Marshal.FreeHGlobal(elevationTypePtr);
                }
            }
            finally
            {
                if (tokenHandle != IntPtr.Zero)
                    CloseHandle(tokenHandle);
            }
        }

        #endregion
    }
}
