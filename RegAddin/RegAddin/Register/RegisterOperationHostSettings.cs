using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RegAddin.Register
{
    [Serializable]
    internal class RegisterOperationHostSettings
    {
        internal RegisterOperationHostSettings(string assemblyPath, SingletonSettings.RegisterMode mode,
            bool codebase, bool doRegisterCall, SingletonSettings.SignCheckMode signCheck,
            SingletonSettings.MetricsMode metrics, SingletonSettings.AddinRegMode addinRegMode)
        {
            AssemblyPath = assemblyPath;
            Mode = mode;
            Codebase = codebase;
            DoRegisterCall = doRegisterCall;
            SignCheck = signCheck;
            Metrics = metrics;
            AddinRegMode = addinRegMode;
        }

        internal string AssemblyPath { get; private set; }

        internal SingletonSettings.RegisterMode Mode { get; private set; }

        internal bool Codebase { get; set; }

        internal bool DoRegisterCall { get; set; }

        internal SingletonSettings.SignCheckMode SignCheck { get; set; }

        internal SingletonSettings.MetricsMode Metrics { get; set; }

        internal SingletonSettings.AddinRegMode AddinRegMode { get; set; }
    }
}
