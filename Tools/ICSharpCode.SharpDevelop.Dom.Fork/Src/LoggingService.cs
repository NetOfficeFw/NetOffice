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
		
		}
		
		public static void Info(object message)
		{
			
		}
		
		public static void Warn(object message)
		{
			
		}
		
		public static void Warn(object message, Exception exception)
		{
			
		}
		
		public static void Error(object message)
		{
			
		}
		
		public static void Error(object message, Exception exception)
		{
			
		}
		
		public static bool IsDebugEnabled
        {
            get
            {
				return false;
			}
		}
	}
}
