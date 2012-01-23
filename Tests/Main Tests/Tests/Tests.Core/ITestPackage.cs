using System;
using System.Collections.Generic;
using System.Text;

namespace Tests.Core
{
    public interface ITestPackage
    {
        /// <summary>
        /// performs a test
        /// </summary>
        /// <param name="logFilePath">folder path for logfile</param>
        /// <returns>test passed or not</returns>
        bool DoTest(string logFilePath);
    }
}
