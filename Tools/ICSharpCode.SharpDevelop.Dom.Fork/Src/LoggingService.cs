// Copyright (c) AlphaSierraPapa for the SharpDevelop Team (for details please see \doc\copyright.txt)
// This code is distributed under the GNU LGPL (for details please see \doc\license.txt)

using System;

namespace ICSharpCode.SharpDevelop.Dom
{
	/// <summary>
	/// We don't reference ICSharpCode.Core but still need the logging interface.
	/// </summary>
	internal static class LoggingService
	{
		public static void Debug(object message)
		{
            if(IsDebugEnabled)
                Console.WriteLine("LoggingService Debug {0}", message);
		}
		
		public static void Info(object message)
		{
            if (IsDebugEnabled)
                Console.WriteLine("LoggingService Info {0}", message);
		}
		
		public static void Warn(object message)
		{
            if (IsDebugEnabled)
                Console.WriteLine("LoggingService Warn {0}", message);
		}
		
		public static void Warn(object message, Exception exception)
		{
            if (IsDebugEnabled)
                Console.WriteLine("LoggingService Warn {0} {1}", message, exception);
		}
		
		public static void Error(object message)
        {
            if (IsDebugEnabled)
                Console.WriteLine("LoggingService Error {0}", message);
		}
		
		public static void Error(object message, Exception exception)
        {
            if (IsDebugEnabled)
                Console.WriteLine("LoggingService Error {0} {1}", message, exception);
		}

        public static bool IsDebugEnabled { get; set; }
       
	}
}
