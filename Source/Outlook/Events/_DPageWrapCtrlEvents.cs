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
    [ComImport, Guid("494F0971-DD96-11D2-AF70-006008AFF117"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _DPageWrapCtrlEvents
	{
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _DPageWrapCtrlEvents_SinkHelper : SinkHelper, _DPageWrapCtrlEvents
	{
		#region Static
		
		public static readonly string Id = "494F0971-DD96-11D2-AF70-006008AFF117";

        #endregion

        #region Ctor

        public _DPageWrapCtrlEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region _DPageWrapCtrlEvents
		
		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}