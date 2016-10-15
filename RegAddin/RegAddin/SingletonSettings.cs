using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin
{
    [Serializable]
    internal class SingletonSettings
    {
        #region Nested
        
        internal enum RegisterMode
        {
            System = 0,
            User = 1
        }

        internal enum UnRegisterMode
        {
            Auto = 0,
            System = 1,
            User = 2
        }

        internal enum ApplicationMode
        {
            Help = 0,
            Register = 1,
            Unregister = 2,
            RegFile = 3,      
        }

        internal enum AlertMode
        {
            Error = 0,
            On = 1,
            Off = 2
        }

        internal enum RegisterCall
        {
            On = 0,
            Off = 1
        }

        internal enum RegExportCall
        {
            On = 0,
            Off = 1
        }

        internal enum SignCheckMode
        {
            Off = 0,
            Warn = 1,
            Error = 2
        }

        internal enum MetricsMode
        {
            None = 0,
            Con = 1,
            Win = 2
        }

        internal enum AddinRegMode
        {
            Off = 0,
            On = 1,
        }

        #endregion

        #region Properties

        internal static ApplicationMode Mode { get; set; }

        internal static RegisterMode RegMode { get; set; }

        internal static UnRegisterMode UnRegMode { get; set; }

        internal static RegExportCall ExportCall { get; set; }

        internal static string AssemblyPath { get; set; }

        internal static string RegFilePath { get; set; }

        internal static AlertMode Alert { get; set; }

        internal static bool Codebase { get; set; }

        internal static RegisterCall DoRegisterCall { get; set; }

        internal static SignCheckMode SignCheck { get; set; }

        internal static bool SuspendMissingAssemblyErrorInUnregister { get; set; } 

        internal static bool Diagnostics { get; set; }

        internal static MetricsMode Metrics { get; set; }

        internal static AddinRegMode AddinReg { get; set; }

        #endregion
    }
}
