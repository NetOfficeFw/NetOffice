using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.PowerPointApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibrary("PP09","PP10","PP11","PP12","PP14")]
	[ComImport, Guid("9149346D-5A91-11CF-8700-00AA0060263B"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface SldEvents
	{
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class SldEvents_SinkHelper : SinkHelper, SldEvents
	{
		#region Static
		
		public static readonly string Id = "9149346D-5A91-11CF-8700-00AA0060263B";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public SldEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region SldEvents Members
		
		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}