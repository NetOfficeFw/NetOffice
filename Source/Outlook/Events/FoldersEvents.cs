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
    [ComImport, Guid("00063076-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface FoldersEvents
	{
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("folder", typeof(NetOffice.OutlookApi.MAPIFolder))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void FolderAdd([In, MarshalAs(UnmanagedType.IDispatch)] object folder);

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("folder", typeof(NetOffice.OutlookApi.MAPIFolder))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void FolderChange([In, MarshalAs(UnmanagedType.IDispatch)] object folder);

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void FolderRemove();
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class FoldersEvents_SinkHelper : SinkHelper, FoldersEvents
	{
		#region Static
		
		public static readonly string Id = "00063076-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

		public FoldersEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region FoldersEvents
		
		public void FolderAdd([In, MarshalAs(UnmanagedType.IDispatch)] object folder)
        {
            if (!Validate("FolderAdd"))
            {
                Invoker.ReleaseParamsArray(folder);
                return;
            }

            NetOffice.OutlookApi.MAPIFolder newFolder = Factory.CreateEventArgumentObjectFromComProxy(EventClass, folder) as NetOffice.OutlookApi.MAPIFolder;
            object[] paramsArray = new object[1];
			paramsArray[0] = newFolder;
			EventBinding.RaiseCustomEvent("FolderAdd", ref paramsArray);
		}

		public void FolderChange([In, MarshalAs(UnmanagedType.IDispatch)] object folder)
		{
            if (!Validate("FolderChange"))
            {
                Invoker.ReleaseParamsArray(folder);
                return;
            }

            NetOffice.OutlookApi.MAPIFolder newFolder = Factory.CreateEventArgumentObjectFromComProxy(EventClass, folder) as NetOffice.OutlookApi.MAPIFolder;
            object[] paramsArray = new object[1];
			paramsArray[0] = newFolder;
			EventBinding.RaiseCustomEvent("FolderChange", ref paramsArray);
		}

		public void FolderRemove()
		{
            if (!Validate("FolderRemove"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("FolderRemove", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}