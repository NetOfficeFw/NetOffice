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
    [ComImport, Guid("0006304F-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ExplorerEvents
	{
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void Activate();

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void FolderSwitch();

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("newFolder", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void BeforeFolderSwitch([In, MarshalAs(UnmanagedType.IDispatch)] object newFolder, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61444)]
		void ViewSwitch();

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61445)]
		void BeforeViewSwitch([In] object newView, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61446)]
		void Deactivate();

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61447)]
		void SelectionChange();

		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61448)]
		void Close();
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ExplorerEvents_SinkHelper : SinkHelper, ExplorerEvents
	{
		#region Static
		
		public static readonly string Id = "0006304F-0000-0000-C000-000000000046";
		
		#endregion

		#region Ctor

		public ExplorerEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region ExplorerEvents
		
		public void Activate()
		{
            if (!Validate("Activate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Activate", ref paramsArray);
		}

		public void FolderSwitch()
		{
            if (!Validate("FolderSwitch"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("FolderSwitch", ref paramsArray);
		}

		public void BeforeFolderSwitch([In, MarshalAs(UnmanagedType.IDispatch)] object newFolder, [In] [Out] ref object cancel)
		{
            if (!Validate("BeforeFolderSwitch"))
            {
                Invoker.ReleaseParamsArray(newFolder, cancel);
                return;
            }

			object newNewFolder = Factory.CreateEventArgumentObjectFromComProxy(EventClass, newFolder) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newNewFolder;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeFolderSwitch", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

		public void ViewSwitch()
		{
            if (!Validate("ViewSwitch"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ViewSwitch", ref paramsArray);
		}

		public void BeforeViewSwitch([In] object newView, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeViewSwitch"))
            {
                Invoker.ReleaseParamsArray(newView, cancel);
                return;
            }

			object newNewView = (object)newView;
			object[] paramsArray = new object[2];
			paramsArray[0] = newNewView;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeViewSwitch", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

		public void Deactivate()
		{
            if (!Validate("Deactivate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Deactivate", ref paramsArray);
		}

		public void SelectionChange()
		{
            if (!Validate("SelectionChange"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("SelectionChange", ref paramsArray);
		}

		public void Close()
		{
            if (!Validate("Close"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Close", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}