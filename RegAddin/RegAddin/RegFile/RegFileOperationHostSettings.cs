using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin.RegFile
{
    [Serializable]
    internal class RegFileOperationHostSettings
    {
        internal RegFileOperationHostSettings(string assemblyPath, SingletonSettings.RegisterMode mode, 
            bool codebase, string regFilePath, SingletonSettings.AddinRegMode addinRegMode, SingletonSettings.RegExportCall exportCall)
        {
            AssemblyPath = assemblyPath;
            Mode = mode;
            Codebase = codebase;
            RegFilePath = regFilePath;
            AddinRegMode = addinRegMode;
            ExportCall = exportCall;
        }
        
        internal string AssemblyPath { get; private set; }

        internal SingletonSettings.RegisterMode Mode { get; private set; }

        internal string RegFilePath { get; private set; }

        internal bool Codebase { get; private set; }

        internal SingletonSettings.AddinRegMode AddinRegMode { get; set; }

        internal SingletonSettings.RegExportCall ExportCall { get; private set; }
    }
}
