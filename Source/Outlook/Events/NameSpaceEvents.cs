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

	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0006308C-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface NameSpaceEvents
	{
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("pages", typeof(OutlookApi.PropertyPages))]
        [SinkArgument("newFolder", typeof(OutlookApi.MAPIFolder))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61445)]
		void OptionsPagesAdd([In, MarshalAs(UnmanagedType.IDispatch)] object pages, [In, MarshalAs(UnmanagedType.IDispatch)] object folder);

		[SupportByVersion("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64557)]
		void AutoDiscoverComplete();
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class NameSpaceEvents_SinkHelper : SinkHelper, NameSpaceEvents
	{
		#region Static
		
		public static readonly string Id = "0006308C-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

		public NameSpaceEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region NameSpaceEvents

        public void OptionsPagesAdd([In, MarshalAs(UnmanagedType.IDispatch)] object pages, [In, MarshalAs(UnmanagedType.IDispatch)] object folder)
		{
            if (!Validate("OptionsPagesAdd"))
            {
                Invoker.ReleaseParamsArray(pages, folder);
                return;
            }

			NetOffice.OutlookApi.PropertyPages newPages = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.PropertyPages>(EventClass, pages, NetOffice.OutlookApi.PropertyPages.LateBindingApiWrapperType);
            NetOffice.OutlookApi.MAPIFolder newFolder = Factory.CreateEventArgumentObjectFromComProxy(EventClass, folder) as NetOffice.OutlookApi.MAPIFolder;

            object[] paramsArray = new object[2];
			paramsArray[0] = newPages;
			paramsArray[1] = newFolder;
			EventBinding.RaiseCustomEvent("OptionsPagesAdd", ref paramsArray);
		}

		public void AutoDiscoverComplete()
        {
            if (!Validate("AutoDiscoverComplete"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("AutoDiscoverComplete", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}