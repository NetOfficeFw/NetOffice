using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.VisioApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Visio", 11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000D0B06-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface EStyle
	{
		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("style", typeof(NetOffice.VisioApi.IVStyle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8196)]
		void StyleChanged([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("style", typeof(NetOffice.VisioApi.IVStyle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16388)]
		void BeforeStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("style", typeof(NetOffice.VisioApi.IVStyle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(300)]
		void QueryCancelStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style);

		[SupportByVersion("Visio", 11,12,14,15,16)]
        [SinkArgument("style", typeof(NetOffice.VisioApi.IVStyle))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(301)]
		void StyleDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object style);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class EStyle_SinkHelper : SinkHelper, EStyle
	{
		#region Static
		
		public static readonly string Id = "000D0B06-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

		public EStyle_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region EStyle
		
		public void StyleChanged([In, MarshalAs(UnmanagedType.IDispatch)] object style)
        {
            if (!Validate("StyleChanged"))
            {
                Invoker.ReleaseParamsArray(style);
                return;
            }

            NetOffice.VisioApi.IVStyle newStyle = Factory.CreateEventArgumentObjectFromComProxy(EventClass, style) as NetOffice.VisioApi.IVStyle;
            object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			EventBinding.RaiseCustomEvent("StyleChanged", ref paramsArray);
		}

		public void BeforeStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
            if (!Validate("BeforeStyleDelete"))
            {
                Invoker.ReleaseParamsArray(style);
                return;
            }

            NetOffice.VisioApi.IVStyle newStyle = Factory.CreateEventArgumentObjectFromComProxy(EventClass, style) as NetOffice.VisioApi.IVStyle;
            object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			EventBinding.RaiseCustomEvent("BeforeStyleDelete", ref paramsArray);
		}

		public void QueryCancelStyleDelete([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
            if (!Validate("QueryCancelStyleDelete"))
            {
                Invoker.ReleaseParamsArray(style);
                return;
            }

            NetOffice.VisioApi.IVStyle newStyle = Factory.CreateEventArgumentObjectFromComProxy(EventClass, style) as NetOffice.VisioApi.IVStyle;
            object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			EventBinding.RaiseCustomEvent("QueryCancelStyleDelete", ref paramsArray);
		}

		public void StyleDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object style)
		{
            if (!Validate("StyleDeleteCanceled"))
            {
                Invoker.ReleaseParamsArray(style);
                return;
            }

            NetOffice.VisioApi.IVStyle newStyle = Factory.CreateEventArgumentObjectFromComProxy(EventClass, style) as NetOffice.VisioApi.IVStyle;
            object[] paramsArray = new object[1];
			paramsArray[0] = newStyle;
			EventBinding.RaiseCustomEvent("StyleDeleteCanceled", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}