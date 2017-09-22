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

	[SupportByVersion("Outlook", 10,11,12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("0006300F-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ExplorerEvents_10
	{
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void Activate();

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void FolderSwitch();

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("newFolder", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void BeforeFolderSwitch([In, MarshalAs(UnmanagedType.IDispatch)] object newFolder, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61444)]
		void ViewSwitch();

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61445)]
		void BeforeViewSwitch([In] object newView, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61446)]
		void Deactivate();

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61447)]
		void SelectionChange();

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61448)]
		void Close();

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64017)]
		void BeforeMaximize([In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64018)]
		void BeforeMinimize([In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64019)]
		void BeforeMove([In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64020)]
		void BeforeSize([In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64014)]
		void BeforeItemCopy([In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64015)]
		void BeforeItemCut([In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("target", typeof(OutlookApi.MAPIFolder))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64016)]
		void BeforeItemPaste([In] [Out] ref object clipboardContent, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64633)]
		void AttachmentSelectionChange();

		[SupportByVersion("Outlook", 15, 16)]
        [SinkArgument("item", SinkArgumentType.UnknownProxy)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64658)]
		void InlineResponse([In, MarshalAs(UnmanagedType.IDispatch)] object item);

        [SupportByVersion("Outlook", 15, 16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64662)]
        void InlineResponseClose();
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ExplorerEvents_10_SinkHelper : SinkHelper, ExplorerEvents_10
	{
		#region Static
		
		public static readonly string Id = "0006300F-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

		public ExplorerEvents_10_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region ExplorerEvents_10
		
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

		public void BeforeMaximize([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeMaximize"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeMaximize", ref paramsArray);

			cancel = ToBoolean(paramsArray[0]);
		}

		public void BeforeMinimize([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeMinimize"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeMinimize", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

		public void BeforeMove([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeMove"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeMove", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

		public void BeforeSize([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeSize"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeSize", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

		public void BeforeItemCopy([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeItemCopy"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeItemCopy", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

		public void BeforeItemCut([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeItemCut"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeItemCut", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

        public void BeforeItemPaste([In] [Out] ref object clipboardContent, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeItemPaste"))
            {
                Invoker.ReleaseParamsArray(clipboardContent, target, cancel);
                return;
            }
            

            NetOffice.OutlookApi.MAPIFolder newTarget = Factory.CreateEventArgumentObjectFromComProxy(EventClass, target) as NetOffice.OutlookApi.MAPIFolder;
			object[] paramsArray = new object[3];
			paramsArray.SetValue(clipboardContent, 0);
			paramsArray[1] = newTarget;
			paramsArray.SetValue(cancel, 2);
			EventBinding.RaiseCustomEvent("BeforeItemPaste", ref paramsArray);

			clipboardContent = (object)paramsArray[0];
			cancel = ToBoolean(paramsArray[2]);
        }

		public void AttachmentSelectionChange()
        {
            if (!Validate("AttachmentSelectionChange"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("AttachmentSelectionChange", ref paramsArray);
		}

		public void InlineResponse([In, MarshalAs(UnmanagedType.IDispatch)] object item)
        {
            if (!Validate("InlineResponse"))
            {
                Invoker.ReleaseParamsArray(item);
                return;
            }

			object newItem = Factory.CreateEventArgumentObjectFromComProxy(EventClass, item) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newItem;
			EventBinding.RaiseCustomEvent("InlineResponse", ref paramsArray);
		}

        public void InlineResponseClose()
        {
            if (!Validate("InlineResponseClose"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("InlineResponseClose", ref paramsArray);
        }

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}