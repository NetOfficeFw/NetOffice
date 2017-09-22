using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Publisher", 14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("00021240-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ApplicationEvents
	{
		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("wn", typeof(PublisherApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("wn", typeof(PublisherApi.Window))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("vw", typeof(PublisherApi.View))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void WindowPageChange([In, MarshalAs(UnmanagedType.IDispatch)] object vw);

		[SupportByVersion("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void Quit();

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void NewDocument([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void DocumentOpen([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void DocumentBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void MailMergeAfterMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
		void MailMergeAfterRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [SinkArgument("startRecord", SinkArgumentType.Int32)]
        [SinkArgument("endRecord", SinkArgumentType.Int32)]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
		void MailMergeBeforeMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object startRecord, [In] object endRecord, [In] [Out] ref object cancel);

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
		void MailMergeBeforeRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void MailMergeDataSourceLoad([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16)]
		void MailMergeWizardSendToCustom([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [SinkArgument("fromState", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(17)]
		void MailMergeWizardStateChange([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object fromState);

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [SinkArgument("handled", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(18)]
		void MailMergeDataSourceValidate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object handled);

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [SinkArgument("okToInsert", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(19)]
		void MailMergeInsertBarcode([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object okToInsert);

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(20)]
		void MailMergeRecipientListClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [SinkArgument("bstrString", SinkArgumentType.String)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(21)]
		void MailMergeGenerateBarcode([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object bstrString);

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(22)]
		void MailMergeWizardFollowUpCustom([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [SinkArgument("cancel", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(23)]
		void BeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		[SupportByVersion("Publisher", 14,15,16)]
        [SinkArgument("doc", typeof(PublisherApi._Document))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(24)]
		void AfterPrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersion("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(25)]
		void ShowCatalogUI();

		[SupportByVersion("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(26)]
		void HideCatalogUI();
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ApplicationEvents_SinkHelper : SinkHelper, ApplicationEvents
	{
		#region Static
		
		public static readonly string Id = "00021240-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

		public ApplicationEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region ApplicationEvents
		
        public void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
            if (!Validate("WindowActivate"))
            {
                Invoker.ReleaseParamsArray(wn);
                return;
            }

			NetOffice.PublisherApi.Window newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PublisherApi.Window>(EventClass, wn, NetOffice.PublisherApi.Window.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			EventBinding.RaiseCustomEvent("WindowActivate", ref paramsArray);
		}

        public void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
            if (!Validate("WindowDeactivate"))
            {
                Invoker.ReleaseParamsArray(wn);
                return;
            }

            NetOffice.PublisherApi.Window newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PublisherApi.Window>(EventClass, wn, NetOffice.PublisherApi.Window.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			EventBinding.RaiseCustomEvent("WindowDeactivate", ref paramsArray);
		}

        public void WindowPageChange([In, MarshalAs(UnmanagedType.IDispatch)] object vw)
        {
            if (!Validate("WindowPageChange"))
            {
                Invoker.ReleaseParamsArray(vw);
                return;
            }

			NetOffice.PublisherApi.View newVw = Factory.CreateKnownObjectFromComProxy<NetOffice.PublisherApi.View>(EventClass, vw, NetOffice.PublisherApi.View.LateBindingApiWrapperType);
			object[] paramsArray = new object[1];
			paramsArray[0] = newVw;
			EventBinding.RaiseCustomEvent("WindowPageChange", ref paramsArray);
		}

		public void Quit()
		{
            if (!Validate("Quit"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Quit", ref paramsArray);
		}

        public void NewDocument([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("NewDocument"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.PublisherApi._Document newDoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.PublisherApi._Document;
            object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			EventBinding.RaiseCustomEvent("NewDocument", ref paramsArray);
		}

        public void DocumentOpen([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
            if (!Validate("DocumentOpen"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.PublisherApi._Document newDoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.PublisherApi._Document;
            object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			EventBinding.RaiseCustomEvent("DocumentOpen", ref paramsArray);
		}

        public void DocumentBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel)
        {
            if (!Validate("DocumentBeforeClose"))
            {
                Invoker.ReleaseParamsArray(doc, cancel);
                return;
            }

            NetOffice.PublisherApi._Document newDoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.PublisherApi._Document;
            object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("DocumentBeforeClose", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

        public void MailMergeAfterMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
            if (!Validate("MailMergeAfterMerge"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.PublisherApi._Document newDoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.PublisherApi._Document;
            object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			EventBinding.RaiseCustomEvent("MailMergeAfterMerge", ref paramsArray);
		}

        public void MailMergeAfterRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
            if (!Validate("MailMergeAfterRecordMerge"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.PublisherApi._Document newDoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.PublisherApi._Document;
            object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			EventBinding.RaiseCustomEvent("MailMergeAfterRecordMerge", ref paramsArray);
		}

        public void MailMergeBeforeMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object startRecord, [In] object endRecord, [In] [Out] ref object cancel)
		{
            if (!Validate("MailMergeBeforeMerge"))
            {
                Invoker.ReleaseParamsArray(doc, startRecord, endRecord, cancel);
                return;
            }

            NetOffice.PublisherApi._Document newDoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.PublisherApi._Document;
            Int32 newStartRecord = ToInt32(startRecord);
			Int32 newEndRecord = ToInt32(endRecord);
			object[] paramsArray = new object[4];
			paramsArray[0] = newDoc;
			paramsArray[1] = newStartRecord;
			paramsArray[2] = newEndRecord;
			paramsArray.SetValue(cancel, 3);
			EventBinding.RaiseCustomEvent("MailMergeBeforeMerge", ref paramsArray);

			cancel = ToBoolean(paramsArray[3]);
		}

        public void MailMergeBeforeRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel)
		{
            if (!Validate("MailMergeBeforeRecordMerge"))
            {
                Invoker.ReleaseParamsArray(doc, cancel);
                return;
            }

            NetOffice.PublisherApi._Document newDoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.PublisherApi._Document;
            object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("MailMergeBeforeRecordMerge", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
        }

        public void MailMergeDataSourceLoad([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("MailMergeDataSourceLoad"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.PublisherApi._Document newDoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.PublisherApi._Document;
            object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			EventBinding.RaiseCustomEvent("MailMergeDataSourceLoad", ref paramsArray);
		}

        public void MailMergeWizardSendToCustom([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("MailMergeWizardSendToCustom"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.PublisherApi._Document newDoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.PublisherApi._Document;
            object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			EventBinding.RaiseCustomEvent("MailMergeWizardSendToCustom", ref paramsArray);
		}

        public void MailMergeWizardStateChange([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object fromState)
        {
            if (!Validate("MailMergeWizardStateChange"))
            {
                Invoker.ReleaseParamsArray(doc, fromState);
                return;
            }

            NetOffice.PublisherApi._Document newDoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.PublisherApi._Document;
            Int32 newFromState = ToInt32(fromState);
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray[1] = newFromState;
			EventBinding.RaiseCustomEvent("MailMergeWizardStateChange", ref paramsArray);
		}

        public void MailMergeDataSourceValidate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object handled)
        {
            if (!Validate("MailMergeDataSourceValidate"))
            {
                Invoker.ReleaseParamsArray(doc, handled);
                return;
            }

            NetOffice.PublisherApi._Document newDoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.PublisherApi._Document;
            object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(handled, 1);
			EventBinding.RaiseCustomEvent("MailMergeDataSourceValidate", ref paramsArray);

			handled = ToBoolean(paramsArray[1]);
		}

        public void MailMergeInsertBarcode([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object okToInsert)
        {
            if (!Validate("MailMergeInsertBarcode"))
            {
                Invoker.ReleaseParamsArray(doc, okToInsert);
                return;
            }

            NetOffice.PublisherApi._Document newDoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.PublisherApi._Document;
            object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(okToInsert, 1);
			EventBinding.RaiseCustomEvent("MailMergeInsertBarcode", ref paramsArray);

			okToInsert = ToBoolean(paramsArray[1]);
        }

        public void MailMergeRecipientListClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("MailMergeRecipientListClose"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.PublisherApi._Document newDoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.PublisherApi._Document;
            object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			EventBinding.RaiseCustomEvent("MailMergeRecipientListClose", ref paramsArray);
		}

        public void MailMergeGenerateBarcode([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object bstrString)
        {
            if (!Validate("MailMergeGenerateBarcode"))
            {
                Invoker.ReleaseParamsArray(doc, bstrString);
                return;
            }

            NetOffice.PublisherApi._Document newDoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.PublisherApi._Document;
            object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(bstrString, 1);
			EventBinding.RaiseCustomEvent("MailMergeGenerateBarcode", ref paramsArray);

			bstrString = ToString(paramsArray[1]);
		}

        public void MailMergeWizardFollowUpCustom([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("MailMergeWizardFollowUpCustom"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.PublisherApi._Document newDoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.PublisherApi._Document;
            object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			EventBinding.RaiseCustomEvent("MailMergeWizardFollowUpCustom", ref paramsArray);
		}

        public void BeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel)
        {
            if (!Validate("BeforePrint"))
            {
                Invoker.ReleaseParamsArray(doc, cancel);
                return;
            }

            NetOffice.PublisherApi._Document newDoc = Factory.CreateEventArgumentObjectFromComProxy(EventClass, doc) as NetOffice.PublisherApi._Document;
            object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(cancel, 1);
			EventBinding.RaiseCustomEvent("BeforePrint", ref paramsArray);

			cancel = ToBoolean(paramsArray[1]);
		}

        public void AfterPrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("AfterPrint"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.PublisherApi._Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.PublisherApi._Document>(EventClass, doc, NetOffice.PublisherApi._Document.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			EventBinding.RaiseCustomEvent("AfterPrint", ref paramsArray);
		}

		public void ShowCatalogUI()
		{
            if (!Validate("ShowCatalogUI"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ShowCatalogUI", ref paramsArray);
		}

		public void HideCatalogUI()
		{
            if (!Validate("HideCatalogUI"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("HideCatalogUI", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}