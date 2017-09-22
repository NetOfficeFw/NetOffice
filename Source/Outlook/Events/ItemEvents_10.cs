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
    [ComImport, Guid("0006302B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ItemEvents_10
	{
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61443)]
		void Open([In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("action", SinkArgumentType.UnknownProxy)]
        [SinkArgument("response", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61446)]
		void CustomAction([In, MarshalAs(UnmanagedType.IDispatch)] object action, [In, MarshalAs(UnmanagedType.IDispatch)] object response, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("name", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61448)]
		void CustomPropertyChange([In] object name);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("forward", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(62568)]
		void Forward([In, MarshalAs(UnmanagedType.IDispatch)] object forward, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61444)]
		void Close([In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("name", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61449)]
		void PropertyChange([In] object name);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61441)]
		void Read();

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("response", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(62566)]
		void Reply([In, MarshalAs(UnmanagedType.IDispatch)] object response, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("response", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(62567)]
		void ReplyAll([In, MarshalAs(UnmanagedType.IDispatch)] object response, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61445)]
		void Send([In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61442)]
		void Write([In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61450)]
		void BeforeCheckNames([In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("attachment", typeof(OutlookApi.Attachment))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61451)]
		void AttachmentAdd([In, MarshalAs(UnmanagedType.IDispatch)] object attachment);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("attachment", typeof(OutlookApi.Attachment))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61452)]
		void AttachmentRead([In, MarshalAs(UnmanagedType.IDispatch)] object attachment);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("attachment", typeof(OutlookApi.Attachment))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(61453)]
		void BeforeAttachmentSave([In, MarshalAs(UnmanagedType.IDispatch)] object attachment, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
        [SinkArgument("item", SinkArgumentType.UnknownProxy)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64117)]
		void BeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("attachment", typeof(OutlookApi.Attachment))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64430)]
		void AttachmentRemove([In, MarshalAs(UnmanagedType.IDispatch)] object attachment);

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("attachment", typeof(OutlookApi.Attachment))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64432)]
		void BeforeAttachmentAdd([In, MarshalAs(UnmanagedType.IDispatch)] object attachment, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("attachment", typeof(OutlookApi.Attachment))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64431)]
		void BeforeAttachmentPreview([In, MarshalAs(UnmanagedType.IDispatch)] object attachment, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("attachment", typeof(OutlookApi.Attachment))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64427)]
		void BeforeAttachmentRead([In, MarshalAs(UnmanagedType.IDispatch)] object attachment, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("attachment", typeof(OutlookApi.Attachment))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64434)]
		void BeforeAttachmentWriteToTempFile([In, MarshalAs(UnmanagedType.IDispatch)] object attachment, [In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64429)]
		void Unload();

		[SupportByVersion("Outlook", 12,14,15,16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64514)]
		void BeforeAutoSave([In] [Out] ref object cancel);

		[SupportByVersion("Outlook", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64652)]
		void BeforeRead();

		[SupportByVersion("Outlook", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64653)]
		void AfterWrite();

		[SupportByVersion("Outlook", 15, 16)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(64655)]
		void ReadComplete([In] [Out] ref object cancel);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class ItemEvents_10_SinkHelper : SinkHelper, ItemEvents_10
    {
        #region Static

        public static readonly string Id = "0006302B-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        public ItemEvents_10_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region ItemEvents_10

        public void Open([In] [Out] ref object cancel)
        {
            if (!Validate("Open"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
            paramsArray.SetValue(cancel, 0);
            EventBinding.RaiseCustomEvent("Open", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

        public void CustomAction([In, MarshalAs(UnmanagedType.IDispatch)] object action, [In, MarshalAs(UnmanagedType.IDispatch)] object response, [In] [Out] ref object cancel)
        {
            if (!Validate("CustomAction"))
            {
                Invoker.ReleaseParamsArray(action, response, cancel);
                return;
            }

            object newAction = Factory.CreateEventArgumentObjectFromComProxy(EventClass, action) as object;
            object newResponse = Factory.CreateEventArgumentObjectFromComProxy(EventClass, response) as object;
            object[] paramsArray = new object[3];
            paramsArray[0] = newAction;
            paramsArray[1] = newResponse;
            paramsArray.SetValue(cancel, 2);
            EventBinding.RaiseCustomEvent("CustomAction", ref paramsArray);

            cancel = ToBoolean(paramsArray[2]);
        }

        public void CustomPropertyChange([In] object name)
        {
            if (!Validate("CustomPropertyChange"))
            {
                Invoker.ReleaseParamsArray(name);
                return;
            }

            string newName = ToString(name);
            object[] paramsArray = new object[1];
            paramsArray[0] = newName;
            EventBinding.RaiseCustomEvent("CustomPropertyChange", ref paramsArray);
        }

        public void Forward([In, MarshalAs(UnmanagedType.IDispatch)] object forward, [In] [Out] ref object cancel)
        {
            if (!Validate("Forward"))
            {
                Invoker.ReleaseParamsArray(forward, cancel);
                return;
            }

            object newForward = Factory.CreateEventArgumentObjectFromComProxy(EventClass, forward) as object;
            object[] paramsArray = new object[2];
            paramsArray[0] = newForward;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("Forward", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void Close([In] [Out] ref object cancel)
        {
            if (!Validate("Close"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
            paramsArray.SetValue(cancel, 0);
            EventBinding.RaiseCustomEvent("Close", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

        public void PropertyChange([In] object name)
        {
            if (!Validate("PropertyChange"))
            {
                Invoker.ReleaseParamsArray(name);
                return;
            }

            string newName = ToString(name);
            object[] paramsArray = new object[1];
            paramsArray[0] = newName;
            EventBinding.RaiseCustomEvent("PropertyChange", ref paramsArray);
        }

        public void Read()
        {
            if (!Validate("Read"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Read", ref paramsArray);
        }

        public void Reply([In, MarshalAs(UnmanagedType.IDispatch)] object response, [In] [Out] ref object cancel)
        {
            if (!Validate("Reply"))
            {
                Invoker.ReleaseParamsArray(response, cancel);
                return;
            }

            object newResponse = Factory.CreateEventArgumentObjectFromComProxy(EventClass, response) as object;
            object[] paramsArray = new object[2];
            paramsArray[0] = newResponse;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("Reply", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void ReplyAll([In, MarshalAs(UnmanagedType.IDispatch)] object response, [In] [Out] ref object cancel)
        {
            if (!Validate("ReplyAll"))
            {
                Invoker.ReleaseParamsArray(response, cancel);
                return;
            }

            object newResponse = Factory.CreateEventArgumentObjectFromComProxy(EventClass, response) as object;
            object[] paramsArray = new object[2];
            paramsArray[0] = newResponse;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("ReplyAll", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void Send([In] [Out] ref object cancel)
        {
            if (!Validate("Send"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
            paramsArray.SetValue(cancel, 0);
            EventBinding.RaiseCustomEvent("Send", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

        public void Write([In] [Out] ref object cancel)
        {
            if (!Validate("Write"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
            paramsArray.SetValue(cancel, 0);
            EventBinding.RaiseCustomEvent("Write", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

        public void BeforeCheckNames([In] [Out] ref object cancel)
        {
            if (!Validate("BeforeCheckNames"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
            paramsArray.SetValue(cancel, 0);
            EventBinding.RaiseCustomEvent("BeforeCheckNames", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

        public void AttachmentAdd([In, MarshalAs(UnmanagedType.IDispatch)] object attachment)
        {
            if (!Validate("AttachmentAdd"))
            {
                Invoker.ReleaseParamsArray(attachment);
                return;
            }

            NetOffice.OutlookApi.Attachment newAttachment = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Attachment>(EventClass, attachment, NetOffice.OutlookApi.Attachment.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newAttachment;
            EventBinding.RaiseCustomEvent("AttachmentAdd", ref paramsArray);
        }

        public void AttachmentRead([In, MarshalAs(UnmanagedType.IDispatch)] object attachment)
        {
            if (!Validate("AttachmentRead"))
            {
                Invoker.ReleaseParamsArray(attachment);
                return;
            }

            NetOffice.OutlookApi.Attachment newAttachment = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Attachment>(EventClass, attachment, NetOffice.OutlookApi.Attachment.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
            paramsArray[0] = newAttachment;
            EventBinding.RaiseCustomEvent("AttachmentRead", ref paramsArray);
        }

        public void BeforeAttachmentSave([In, MarshalAs(UnmanagedType.IDispatch)] object attachment, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeAttachmentSave"))
            {
                Invoker.ReleaseParamsArray(attachment, cancel);
                return;
            }

            NetOffice.OutlookApi.Attachment newAttachment = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Attachment>(EventClass, attachment, NetOffice.OutlookApi.Attachment.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newAttachment;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeAttachmentSave", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

        public void BeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object item, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeDelete"))
            {
                Invoker.ReleaseParamsArray(item, cancel);
                return;
            }

			object newItem = Factory.CreateEventArgumentObjectFromComProxy(EventClass, item) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newItem;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeDelete", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void AttachmentRemove([In, MarshalAs(UnmanagedType.IDispatch)] object attachment)
		{
            if (!Validate("AttachmentRemove"))
            {
                Invoker.ReleaseParamsArray(attachment);
                return;
            }

            NetOffice.OutlookApi.Attachment newAttachment = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Attachment>(EventClass, attachment, NetOffice.OutlookApi.Attachment.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newAttachment;
			EventBinding.RaiseCustomEvent("AttachmentRemove", ref paramsArray);
		}

        public void BeforeAttachmentAdd([In, MarshalAs(UnmanagedType.IDispatch)] object attachment, [In] [Out] ref object cancel)
		{
            if (!Validate("BeforeAttachmentAdd"))
            {
                Invoker.ReleaseParamsArray(attachment, cancel);
                return;
            }

            NetOffice.OutlookApi.Attachment newAttachment = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Attachment>(EventClass, attachment, NetOffice.OutlookApi.Attachment.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newAttachment;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeAttachmentAdd", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

        public void BeforeAttachmentPreview([In, MarshalAs(UnmanagedType.IDispatch)] object attachment, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeAttachmentPreview"))
            {
                Invoker.ReleaseParamsArray(attachment, cancel);
                return;
            }

            NetOffice.OutlookApi.Attachment newAttachment = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Attachment>(EventClass, attachment, NetOffice.OutlookApi.Attachment.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newAttachment;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeAttachmentPreview", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void BeforeAttachmentRead([In, MarshalAs(UnmanagedType.IDispatch)] object attachment, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforeAttachmentRead"))
            {
                Invoker.ReleaseParamsArray(attachment, cancel);
                return;
            }

            NetOffice.OutlookApi.Attachment newAttachment = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Attachment>(EventClass, attachment, NetOffice.OutlookApi.Attachment.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newAttachment;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeAttachmentRead", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void BeforeAttachmentWriteToTempFile([In, MarshalAs(UnmanagedType.IDispatch)] object attachment, [In] [Out] ref object cancel)
		{
            if (!Validate("BeforeAttachmentWriteToTempFile"))
            {
                Invoker.ReleaseParamsArray(attachment, cancel);
                return;
            }

            NetOffice.OutlookApi.Attachment newAttachment = Factory.CreateKnownObjectFromComProxy<NetOffice.OutlookApi.Attachment>(EventClass, attachment, NetOffice.OutlookApi.Attachment.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
			paramsArray[0] = newAttachment;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforeAttachmentWriteToTempFile", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        public void Unload()
		{
            if (!Validate("Unload"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Unload", ref paramsArray);
		}

        public void BeforeAutoSave([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeAutoSave"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeAutoSave", ref paramsArray);

			cancel = ToBoolean(paramsArray[0]);
		}

		public void BeforeRead()
		{
            if (!Validate("BeforeRead"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("BeforeRead", ref paramsArray);
		}

		public void AfterWrite()
		{
            if (!Validate("AfterWrite"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("AfterWrite", ref paramsArray);
		}

        public void ReadComplete([In] [Out] ref object cancel)
		{
            if (!Validate("ReadComplete"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("ReadComplete", ref paramsArray);

			cancel = ToBoolean(paramsArray[0]);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}