using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Outlook", 10)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("2D9C6D57-BD3C-4275-BED2-73F0EDC18CCE"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _DInspectorEvents
	{
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _DInspectorEvents_SinkHelper : SinkHelper, _DInspectorEvents
	{
		#region Static
		
		public static readonly string Id = "2D9C6D57-BD3C-4275-BED2-73F0EDC18CCE";
		
		#endregion
	
		#region Ctor

		public _DInspectorEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region _DInspectorEvents Members
		
		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}