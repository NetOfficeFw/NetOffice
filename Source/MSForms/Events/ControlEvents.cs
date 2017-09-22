using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("MSForms", 2)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("9A4BBF53-4E46-101B-8BBD-00AA003E3B29"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ControlEvents
	{
		[SupportByVersion("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147384830)]
		void Enter();

		[SupportByVersion("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147384829)]
		void Exit([In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("cancel", typeof(NetOffice.MSFormsApi.ReturnBoolean))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147384831)]
		void BeforeUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersion("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-2147384832)]
		void AfterUpdate();
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ControlEvents_SinkHelper : SinkHelper, ControlEvents
	{
		#region Static
		
		public static readonly string Id = "9A4BBF53-4E46-101B-8BBD-00AA003E3B29";
		
		#endregion
		
		#region Ctor

		public ControlEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region ControlEvents
		
		public void Enter()
        {
            if (!Validate("Enter"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Enter", ref paramsArray);
		}

		public void Exit([In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
        {
            if (!Validate("Exit"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			NetOffice.MSFormsApi.ReturnBoolean newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnBoolean>(EventClass, cancel, NetOffice.MSFormsApi.ReturnBoolean.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newCancel;
			EventBinding.RaiseCustomEvent("Exit", ref paramsArray);
		}

		public void BeforeUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
        {
            if (!Validate("BeforeUpdate"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            NetOffice.MSFormsApi.ReturnBoolean newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnBoolean>(EventClass, cancel, NetOffice.MSFormsApi.ReturnBoolean.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newCancel;
			EventBinding.RaiseCustomEvent("BeforeUpdate", ref paramsArray);
		}

		public void AfterUpdate()
        {
            if (!Validate("AfterUpdate"))
            {   
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("AfterUpdate", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}