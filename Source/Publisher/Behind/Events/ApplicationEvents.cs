using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice.PublisherApi.Behind.EventContracts
{
	/// <summary>
	/// Default implementation of <see cref="NetOffice.PublisherApi.EventContracts.ApplicationEvents"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ApplicationEvents_SinkHelper : SinkHelper, NetOffice.PublisherApi.EventContracts.ApplicationEvents
	{
		#region Static
		
		/// <summary>
		/// Interface Id from ApplicationEvents
		/// </summary>
		public static readonly string Id = "00021240-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public ApplicationEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region ApplicationEvents
		
		/// <summary>
		/// 
		/// </summary>
		/// <param name="wn"></param>
        public void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
            if (!Validate("WindowActivate"))
            {
                Invoker.ReleaseParamsArray(wn);
                return;
            }

			NetOffice.PublisherApi.Window newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PublisherApi.Window>(EventClass, wn, typeof(NetOffice.PublisherApi.Window));
			object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			EventBinding.RaiseCustomEvent("WindowActivate", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="wn"></param>
        public void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
            if (!Validate("WindowDeactivate"))
            {
                Invoker.ReleaseParamsArray(wn);
                return;
            }

            NetOffice.PublisherApi.Window newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.PublisherApi.Window>(EventClass, wn, typeof(NetOffice.PublisherApi.Window));
            object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			EventBinding.RaiseCustomEvent("WindowDeactivate", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="vw"></param>
        public void WindowPageChange([In, MarshalAs(UnmanagedType.IDispatch)] object vw)
        {
            if (!Validate("WindowPageChange"))
            {
                Invoker.ReleaseParamsArray(vw);
                return;
            }

			NetOffice.PublisherApi.View newVw = Factory.CreateKnownObjectFromComProxy<NetOffice.PublisherApi.View>(EventClass, vw, typeof(NetOffice.PublisherApi.View));
			object[] paramsArray = new object[1];
			paramsArray[0] = newVw;
			EventBinding.RaiseCustomEvent("WindowPageChange", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void Quit()
		{
            if (!Validate("Quit"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Quit", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="cancel"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="startRecord"></param>
		/// <param name="endRecord"></param>
		/// <param name="cancel"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="cancel"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="fromState"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="handled"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="okToInsert"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="bstrString"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="cancel"></param>
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="doc"></param>
        public void AfterPrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("AfterPrint"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.PublisherApi._Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.PublisherApi._Document>(EventClass, doc, typeof(NetOffice.PublisherApi._Document));
            object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			EventBinding.RaiseCustomEvent("AfterPrint", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void ShowCatalogUI()
		{
            if (!Validate("ShowCatalogUI"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("ShowCatalogUI", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
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
}
