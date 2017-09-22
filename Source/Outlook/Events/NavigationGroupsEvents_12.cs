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

	[SupportByVersion("Outlook", 12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000630F4-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface NavigationGroupsEvents_12
	{
		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("navigationFolder", typeof(OutlookApi.NavigationFolder))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64458)]
		void SelectedChange([In, MarshalAs(UnmanagedType.IDispatch)] object navigationFolder);

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("navigationFolder", typeof(OutlookApi.NavigationFolder))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64459)]
		void NavigationFolderAdd([In, MarshalAs(UnmanagedType.IDispatch)] object navigationFolder);

		[SupportByVersion("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64460)]
		void NavigationFolderRemove();
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class NavigationGroupsEvents_12_SinkHelper : SinkHelper, NavigationGroupsEvents_12
	{
		#region Static
		
		public static readonly string Id = "000630F4-0000-0000-C000-000000000046";
		
		#endregion

		#region Ctor

		public NavigationGroupsEvents_12_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region NavigationGroupsEvents_12
		
		public void SelectedChange([In, MarshalAs(UnmanagedType.IDispatch)] object navigationFolder)
		{
            if (!Validate("SelectedChange"))
            {
                Invoker.ReleaseParamsArray(navigationFolder);
                return;
            }

			NetOffice.OutlookApi.NavigationFolder newNavigationFolder = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.NavigationFolder>(EventClass, navigationFolder, NetOffice.OutlookApi.NavigationFolder.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newNavigationFolder;
			EventBinding.RaiseCustomEvent("SelectedChange", ref paramsArray);
		}

		public void NavigationFolderAdd([In, MarshalAs(UnmanagedType.IDispatch)] object navigationFolder)
		{
            if (!Validate("NavigationFolderAdd"))
            {
                Invoker.ReleaseParamsArray(navigationFolder);
                return;
            }

            NetOffice.OutlookApi.NavigationFolder newNavigationFolder = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.NavigationFolder>(EventClass, navigationFolder, NetOffice.OutlookApi.NavigationFolder.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newNavigationFolder;
			EventBinding.RaiseCustomEvent("NavigationFolderAdd", ref paramsArray);
		}

		public void NavigationFolderRemove()
        {
            if (!Validate("NavigationFolderRemove"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("NavigationFolderRemove", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}