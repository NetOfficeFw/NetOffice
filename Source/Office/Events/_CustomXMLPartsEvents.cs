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

	[SupportByVersion("Office", 12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000CDB0B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CustomXMLPartsEvents
	{
		[SupportByVersion("Office", 12,14,15,16)]
        [SinkArgument("newPart", typeof(OfficeApi.CustomXMLPart))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void PartAfterAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newPart);

		[SupportByVersion("Office", 12,14,15,16)]
        [SinkArgument("oldPart", typeof(OfficeApi.CustomXMLPart))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void PartBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object oldPart);

		[SupportByVersion("Office", 12,14,15,16)]
        [SinkArgument("part", typeof(OfficeApi.CustomXMLPart))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void PartAfterLoad([In, MarshalAs(UnmanagedType.IDispatch)] object part);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _CustomXMLPartsEvents_SinkHelper : SinkHelper, _CustomXMLPartsEvents
	{
		#region Static
		
		public static readonly string Id = "000CDB0B-0000-0000-C000-000000000046";
		
		#endregion
		
		#region Ctor

		public _CustomXMLPartsEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region _CustomXMLPartsEvents

        public void PartAfterAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newPart)
        {
            if (!Validate("PartAfterAdd"))
            {
                Invoker.ReleaseParamsArray(newPart);
                return;
            }
            
			NetOffice.OfficeApi.CustomXMLPart newNewPart = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CustomXMLPart>(EventClass, newPart, NetOffice.OfficeApi.CustomXMLPart.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newNewPart;
			EventBinding.RaiseCustomEvent("PartAfterAdd", ref paramsArray);
		}

        public void PartBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object oldPart)
        {
            if (!Validate("PartBeforeDelete"))
            {
                Invoker.ReleaseParamsArray(oldPart);
                return;
            }

            NetOffice.OfficeApi.CustomXMLPart newOldPart = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CustomXMLPart>(EventClass, oldPart, NetOffice.OfficeApi.CustomXMLPart.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newOldPart;
			EventBinding.RaiseCustomEvent("PartBeforeDelete", ref paramsArray);
		}
        
        public void PartAfterLoad([In, MarshalAs(UnmanagedType.IDispatch)] object part)
		{
            if (!Validate("PartAfterLoad"))
            {
                Invoker.ReleaseParamsArray(part);
                return;
            }

            NetOffice.OfficeApi.CustomXMLPart newPart = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CustomXMLPart>(EventClass, part, NetOffice.OfficeApi.CustomXMLPart.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newPart;
			EventBinding.RaiseCustomEvent("PartAfterLoad", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}