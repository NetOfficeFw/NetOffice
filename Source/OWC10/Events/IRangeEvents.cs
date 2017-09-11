using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("OWC10", 1)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("B8891063-2B00-48EC-957F-6DEBEADE9D8B"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface IRangeEvents
	{
		[SupportByVersion("OWC10", 1)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1510)]
		void Change();
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class IRangeEvents_SinkHelper : SinkHelper, IRangeEvents
	{
		#region Static
		
		public static readonly string Id = "B8891063-2B00-48EC-957F-6DEBEADE9D8B";
		
		#endregion

		#region Ctor

		public IRangeEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region IRangeEvents
		
		public void Change()
		{
            if (!Validate("Change"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Change", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}