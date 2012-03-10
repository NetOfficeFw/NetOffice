using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Tests.Core
{
    public delegate void OnErrorEventHandler(ITestPackage sender, Exception exception);

    public interface ITestAssembly
    {
        /// <summary>
        /// used programming language
        /// </summary>
        string Language { get; }

        /// <summary>
        /// target office product
        /// </summary>
        string OfficeProduct { get; }

        /// <summary>
        /// load the tests
        /// </summary>
        /// <returns></returns>
        ITestPackage[] LoadTestPackages();
    }
}
