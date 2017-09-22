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
    [ComImport, Guid("4BD09D02-45CC-11D1-B1D1-006097C97F9B"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _NavigationEvent
	{
		[SupportByVersion("OWC10", 1)]
        [SinkArgument("navButton", typeof(OWC10Api.Enums.NavButtonEnum))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(740)]
		void ButtonClick([In] object navButton);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _NavigationEvent_SinkHelper : SinkHelper, _NavigationEvent
	{
		#region Static
		
		public static readonly string Id = "4BD09D02-45CC-11D1-B1D1-006097C97F9B";
		
		#endregion
	
		#region Ctor

		public _NavigationEvent_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region _NavigationEvent
		
		public void ButtonClick([In] object navButton)
		{
            if (!Validate("ButtonClick"))
            {
                Invoker.ReleaseParamsArray(navButton);
                return;
            }

			NetOffice.OWC10Api.Enums.NavButtonEnum newNavButton = (NetOffice.OWC10Api.Enums.NavButtonEnum)navButton;
			object[] paramsArray = new object[1];
			paramsArray[0] = newNavButton;
			EventBinding.RaiseCustomEvent("ButtonClick", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}