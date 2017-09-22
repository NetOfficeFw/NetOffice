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

	[SupportByVersion("Outlook", 10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000630A5-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _ViewsEvents
	{
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("view", typeof(OutlookApi.View))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(400)]
		void ViewAdd([In, MarshalAs(UnmanagedType.IDispatch)] object view);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("view", typeof(OutlookApi.View))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64071)]
		void ViewRemove([In, MarshalAs(UnmanagedType.IDispatch)] object view);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _ViewsEvents_SinkHelper : SinkHelper, _ViewsEvents
	{
		#region Static
		
		public static readonly string Id = "000630A5-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        public _ViewsEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}

        #endregion

        #region _ViewsEvents

        public void ViewAdd([In, MarshalAs(UnmanagedType.IDispatch)] object view)
        {
            if (!Validate("ViewAdd"))
            {
                Invoker.ReleaseParamsArray(view);
            }

			NetOffice.OutlookApi.View newView = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.View>(EventClass, view, NetOffice.OutlookApi.View.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newView;
			EventBinding.RaiseCustomEvent("ViewAdd", ref paramsArray);
		}

		public void ViewRemove([In, MarshalAs(UnmanagedType.IDispatch)] object view)
        {
            if (!Validate("ViewRemove"))
            {
                Invoker.ReleaseParamsArray(view);
            }

            NetOffice.OutlookApi.View newView = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.View>(EventClass, view, NetOffice.OutlookApi.View.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newView;
			EventBinding.RaiseCustomEvent("ViewRemove", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}