using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Tests.Core
{
    public class TestResult
    {
        public TestResult(bool sucseed, TimeSpan timeElapsed, string errorInfo, Exception exception, string hints)
        {
            Sucseed = sucseed;
            TimeElapsed = timeElapsed;
            ErrorInfo = errorInfo;
            Exception = exception;
            Hints = hints;
        }

        public bool Sucseed { get; internal set; }
        public TimeSpan TimeElapsed { get; internal set; }
        public string ErrorInfo { get; internal set; }
        public Exception Exception { get; internal set; }
        public string Hints { get; internal set; }
    }
}
