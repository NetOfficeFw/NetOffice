using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi.EventInterfaces
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("VBIDE", 12,14,5.3)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("CDDE3804-2064-11CF-867F-00AA005FF34A"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _dispReferences_Events
	{
		[SupportByVersion("VBIDE", 12,14,5.3)]
        [SinkArgument("reference", typeof(NetOffice.VBIDEApi.Reference))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0)]
		void ItemAdded([In, MarshalAs(UnmanagedType.IDispatch)] object reference);

		[SupportByVersion("VBIDE", 12,14,5.3)]
        [SinkArgument("reference", typeof(NetOffice.VBIDEApi.Reference))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void ItemRemoved([In, MarshalAs(UnmanagedType.IDispatch)] object reference);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _dispReferences_Events_SinkHelper : SinkHelper, _dispReferences_Events
	{
		#region Static
		
		public static readonly string Id = "CDDE3804-2064-11CF-867F-00AA005FF34A";
		
		#endregion
		
		#region Ctor

		public _dispReferences_Events_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region _dispReferences_Events
		
		public void ItemAdded([In, MarshalAs(UnmanagedType.IDispatch)] object reference)
        {
            if (!Validate("ItemAdded"))
            {
                Invoker.ReleaseParamsArray(reference);
                return;
            }

			NetOffice.VBIDEApi.Reference newReference = Factory.CreateKnownObjectFromComProxy<NetOffice.VBIDEApi.Reference>(EventClass, reference, NetOffice.VBIDEApi.Reference.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newReference;
			EventBinding.RaiseCustomEvent("ItemAdded", ref paramsArray);
		}

		public void ItemRemoved([In, MarshalAs(UnmanagedType.IDispatch)] object reference)
        {
            if (!Validate("ItemRemoved"))
            {
                Invoker.ReleaseParamsArray(reference);
                return;
            }

            NetOffice.VBIDEApi.Reference newReference = Factory.CreateKnownObjectFromComProxy<NetOffice.VBIDEApi.Reference>(EventClass, reference, NetOffice.VBIDEApi.Reference.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newReference;
            EventBinding.RaiseCustomEvent("ItemRemoved", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}