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

	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("50BB9B50-811D-11CE-B565-00AA00608FAA"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _DDocSiteControlEvents
	{
	}
	
	#endregion
	
	#region SinkHelper
	
    [InternalEntity(InternalEntityKind.SinkHelper)]
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _DDocSiteControlEvents_SinkHelper : SinkHelper, _DDocSiteControlEvents
	{
		#region Static
		
		public static readonly string Id = "50BB9B50-811D-11CE-B565-00AA00608FAA";
		
		#endregion
		
		#region Ctor

		public _DDocSiteControlEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region _DDocSiteControlEvents Members
		
		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}