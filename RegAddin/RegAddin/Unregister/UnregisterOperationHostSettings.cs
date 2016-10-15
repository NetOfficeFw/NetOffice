using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin.Unregister
{
    [Serializable]
    internal class UnregisterOperationHostSettings
    {
        internal UnregisterOperationHostSettings(string assemblyPath, SingletonSettings.UnRegisterMode mode,
            bool doRegisterCall, SingletonSettings.SignCheckMode signCheck,
            SingletonSettings.AddinRegMode addinRegMode)
        {
            AssemblyPath = assemblyPath;
            Mode = mode;
            DoRegisterCall = doRegisterCall;
            SignCheck = signCheck;
            AddinRegMode = addinRegMode;
        }

        internal string AssemblyPath { get; private set; }

        internal SingletonSettings.UnRegisterMode Mode { get; private set; }

        internal bool DoRegisterCall { get; private set; }

        internal SingletonSettings.SignCheckMode SignCheck { get; private set; }

        internal SingletonSettings.AddinRegMode AddinRegMode { get; set; }
    }
}
