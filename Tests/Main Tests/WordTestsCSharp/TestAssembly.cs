using System;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tests.Core;

namespace WordTestsCSharp
{
    public class TestAssembly : ITestAssembly
    {
        private List<ITestPackage> _listPackages;

        internal static TestAssembly Singleton { get; private set; }

        public TestAssembly()
        {
            Singleton = this;
        }

        #region ITestAssembly Members

        public string Language
        {
            get { return "C#"; }
        }

        public string OfficeProduct
        {
            get { return "Word"; }
        }

        public ITestPackage[] LoadTestPackages()
        {
            if (null == _listPackages)
            {
                NetOffice.DebugConsole.Default.Mode = NetOffice.DebugConsoleMode.Console;
                NetOffice.DebugConsole.Default.EnableSharedOutput = true;

                AddRegistryTweaks();

                _listPackages = new List<ITestPackage>();
                _listPackages.Add(new Test01());
                _listPackages.Add(new Test02());
                _listPackages.Add(new Test03());
                _listPackages.Add(new Test04());
                _listPackages.Add(new Test05());
                _listPackages.Add(new Test06());
                _listPackages.Add(new Test07());
                _listPackages.Add(new Test08());
                _listPackages.Add(new Test09());
            }
            return _listPackages.ToArray();
        }

        #endregion

        private void AddRegistryTweaks()
        {
            RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\Microsoft\\Office\\Word\\Addins\\NOTestsMain.WordTestAddinCSharp", true);
            if (null != key)
            {
                key.SetValue("NOExceptionMessage", "WordTweakCS", RegistryValueKind.String);
                key.Close();
                key.Dispose();
            }
        }
    }
}
