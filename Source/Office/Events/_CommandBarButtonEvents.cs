using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000C0351-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CommandBarButtonEvents
	{
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
        [SinkArgument("ctrl", typeof(CommandBarButton))]
        [SinkArgument("cancelDefault", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Click([In, MarshalAs(UnmanagedType.IDispatch)] object ctrl, [In] [Out] ref object cancelDefault);
	}
	
	#endregion
	
	#region SinkHelper
	
    [InternalEntity(InternalEntityKind.SinkHelper)]
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _CommandBarButtonEvents_SinkHelper : SinkHelper, _CommandBarButtonEvents
	{
		#region Static
		
		public static readonly string Id = "000C0351-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        public _CommandBarButtonEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region _CommandBarButtonEvents
		
		public void Click([In, MarshalAs(UnmanagedType.IDispatch)] object ctrl, [In] [Out] ref object cancelDefault)
		{
            if(!Validate("Click"))
            {
                Invoker.ReleaseParamsArray(ctrl, cancelDefault);
                return;
            }

            NetOffice.OfficeApi.CommandBarButton newCtrl = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CommandBarButton>(EventClass, ctrl, NetOffice.OfficeApi.CommandBarButton.LateBindingApiWrapperType);
			object[] paramsArray = new object[2];
			paramsArray[0] = newCtrl;
			paramsArray.SetValue(cancelDefault, 1);
			EventBinding.RaiseCustomEvent("Click", ref paramsArray);

            cancelDefault = ToBoolean(paramsArray[1]);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}