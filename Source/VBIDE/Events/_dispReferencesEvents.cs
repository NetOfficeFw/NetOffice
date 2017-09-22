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
    [ComImport, Guid("0002E118-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _dispReferencesEvents
	{
		[SupportByVersion("VBIDE", 12,14,5.3)]
        [SinkArgument("reference", typeof(NetOffice.VBIDEApi.Reference))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void ItemAdded([In, MarshalAs(UnmanagedType.IDispatch)] object reference);

		[SupportByVersion("VBIDE", 12,14,5.3)]
        [SinkArgument("reference", typeof(NetOffice.VBIDEApi.Reference))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void ItemRemoved([In, MarshalAs(UnmanagedType.IDispatch)] object reference);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _dispReferencesEvents_SinkHelper : SinkHelper, _dispReferencesEvents
	{
		#region Static
		
		public static readonly string Id = "0002E118-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

		public _dispReferencesEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region _dispReferencesEvents
		
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