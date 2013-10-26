using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tests.Core;

namespace ProjectTestsCSharp
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
            get { return "Project"; }
        }

        public ITestPackage[] LoadTestPackages()
        {
            if (null == _listPackages)
            {
                _listPackages = new List<ITestPackage>();
                _listPackages.Add(new Test01());
                _listPackages.Add(new Test02());

            }
            return _listPackages.ToArray();
        }

        #endregion
    }
}
