using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace RegAddin
{
    internal class VersionPresenter
    {
        internal void ShowVersion()
        {
            Assembly assembly = typeof(VersionPresenter).Assembly;
            AssemblyName assemblyName = assembly.GetName();
            Version version = assembly.GetName().Version;
            Console.WriteLine("{0} [Version {1}.{2}]", assembly.GetName().Name, version.Major, version.Minor);
        }
    }
}
