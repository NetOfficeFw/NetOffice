using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.AccessApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("F163F201-ADA2-11CF-89A9-00A0C9054129"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _References_Events
	{
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("reference", typeof(NetOffice.AccessApi.Reference))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0)]
		void ItemAdded([In, MarshalAs(UnmanagedType.IDispatch)] object reference);

		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
        [SinkArgument("reference", typeof(NetOffice.AccessApi.Reference))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void ItemRemoved([In, MarshalAs(UnmanagedType.IDispatch)] object reference);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _References_Events_SinkHelper : SinkHelper, _References_Events
	{
		#region Static
		
		public static readonly string Id = "F163F201-ADA2-11CF-89A9-00A0C9054129";
		
		#endregion
		
		#region Ctor

		public _References_Events_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region _References_Events
		
		public void ItemAdded([In, MarshalAs(UnmanagedType.IDispatch)] object reference)
		{
            if (!Validate("ItemAdded"))
            {
                Invoker.ReleaseParamsArray(reference);
                return;
            }

			NetOffice.AccessApi.Reference newReference = Factory.CreateKnownObjectFromComProxy<NetOffice.AccessApi.Reference>(EventClass, reference, NetOffice.AccessApi.Reference.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newReference;
			EventBinding.RaiseCustomEvent("ItemAdded", ref paramsArray);
		}

		public void ItemRemoved([In, MarshalAs(UnmanagedType.IDispatch)] object reference)
        {
            if (!Validate("ItemAdded"))
            {
                Invoker.ReleaseParamsArray(reference);
                return;
            }

            NetOffice.AccessApi.Reference newReference = Factory.CreateKnownObjectFromComProxy<NetOffice.AccessApi.Reference>(EventClass, reference, NetOffice.AccessApi.Reference.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newReference;
			EventBinding.RaiseCustomEvent("ItemRemoved", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}