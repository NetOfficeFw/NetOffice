using System;
using System.Collections.Generic;
using System.Text;

namespace Tests.Core
{
    public interface ITestPackage
    {
        /// <summary>
        /// name of the test
        /// </summary>
        string Name { get; }

        /// <summary>
        /// description of the test
        /// </summary>
        string Description { get; }

        /// <summary>
        /// target office product
        /// </summary>
        string OfficeProduct { get; }

        /// <summary>
        /// used programming language
        /// </summary>
        string Language { get; }

        /// <summary>
        /// performs the test
        /// </summary>
        /// <returns></returns>
        TestResult DoTest();
    }
}
