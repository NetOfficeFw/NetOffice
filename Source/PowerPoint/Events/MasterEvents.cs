using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("914934D2-5A91-11CF-8700-00AA0060263B"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface MasterEvents
	{
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class MasterEvents_SinkHelper : SinkHelper, MasterEvents
	{
		#region Static
		
		public static readonly string Id = "914934D2-5A91-11CF-8700-00AA0060263B";
		
		#endregion
	
		#region Ctor

		public MasterEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
        {
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region MasterEvents Members
		
		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}