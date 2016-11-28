using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.PublisherApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("Publisher", 14,15,16)]
	[ComImport, Guid("00021240-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ApplicationEvents
	{
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void WindowPageChange([In, MarshalAs(UnmanagedType.IDispatch)] object vw);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
		void Quit();

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void NewDocument([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void DocumentOpen([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void DocumentBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void MailMergeAfterMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
		void MailMergeAfterRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
		void MailMergeBeforeMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object startRecord, [In] object endRecord, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
		void MailMergeBeforeRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void MailMergeDataSourceLoad([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16)]
		void MailMergeWizardSendToCustom([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(17)]
		void MailMergeWizardStateChange([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object fromState);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(18)]
		void MailMergeDataSourceValidate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object handled);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(19)]
		void MailMergeInsertBarcode([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object okToInsert);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(20)]
		void MailMergeRecipientListClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(21)]
		void MailMergeGenerateBarcode([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object bstrString);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(22)]
		void MailMergeWizardFollowUpCustom([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(23)]
		void BeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(24)]
		void AfterPrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(25)]
		void ShowCatalogUI();

		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(26)]
		void HideCatalogUI();
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ApplicationEvents_SinkHelper : SinkHelper, ApplicationEvents
	{
		#region Static
		
		public static readonly string Id = "00021240-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private ICOMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public ApplicationEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region Properties

        internal Core Factory
        {
            get
            {
                if (null != _eventClass)
                    return _eventClass.Factory;
                else
                    return Core.Default;
            }
        }

        #endregion

		#region ApplicationEvents Members
		
		public void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowActivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wn);
				return;
			}

			NetOffice.PublisherApi.Window newWn = Factory.CreateObjectFromComProxy(_eventClass, wn) as NetOffice.PublisherApi.Window;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			_eventBinding.RaiseCustomEvent("WindowActivate", ref paramsArray);
		}

		public void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowDeactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wn);
				return;
			}

			NetOffice.PublisherApi.Window newWn = Factory.CreateObjectFromComProxy(_eventClass, wn) as NetOffice.PublisherApi.Window;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			_eventBinding.RaiseCustomEvent("WindowDeactivate", ref paramsArray);
		}

		public void WindowPageChange([In, MarshalAs(UnmanagedType.IDispatch)] object vw)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowPageChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(vw);
				return;
			}

			NetOffice.PublisherApi.View newVw = Factory.CreateObjectFromComProxy(_eventClass, vw) as NetOffice.PublisherApi.View;
			object[] paramsArray = new object[1];
			paramsArray[0] = newVw;
			_eventBinding.RaiseCustomEvent("WindowPageChange", ref paramsArray);
		}

		public void Quit()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Quit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("Quit", ref paramsArray);
		}

		public void NewDocument([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("NewDocument");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.PublisherApi._Document newDoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.PublisherApi._Document;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			_eventBinding.RaiseCustomEvent("NewDocument", ref paramsArray);
		}

		public void DocumentOpen([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DocumentOpen");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.PublisherApi._Document newDoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.PublisherApi._Document;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			_eventBinding.RaiseCustomEvent("DocumentOpen", ref paramsArray);
		}

		public void DocumentBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DocumentBeforeClose");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, cancel);
				return;
			}

			NetOffice.PublisherApi._Document newDoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.PublisherApi._Document;
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("DocumentBeforeClose", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void MailMergeAfterMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MailMergeAfterMerge");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.PublisherApi._Document newDoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.PublisherApi._Document;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			_eventBinding.RaiseCustomEvent("MailMergeAfterMerge", ref paramsArray);
		}

		public void MailMergeAfterRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MailMergeAfterRecordMerge");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.PublisherApi._Document newDoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.PublisherApi._Document;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			_eventBinding.RaiseCustomEvent("MailMergeAfterRecordMerge", ref paramsArray);
		}

		public void MailMergeBeforeMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object startRecord, [In] object endRecord, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MailMergeBeforeMerge");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, startRecord, endRecord, cancel);
				return;
			}

			NetOffice.PublisherApi._Document newDoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.PublisherApi._Document;
			Int32 newStartRecord = Convert.ToInt32(startRecord);
			Int32 newEndRecord = Convert.ToInt32(endRecord);
			object[] paramsArray = new object[4];
			paramsArray[0] = newDoc;
			paramsArray[1] = newStartRecord;
			paramsArray[2] = newEndRecord;
			paramsArray.SetValue(cancel, 3);
			_eventBinding.RaiseCustomEvent("MailMergeBeforeMerge", ref paramsArray);

			cancel = (bool)paramsArray[3];
		}

		public void MailMergeBeforeRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MailMergeBeforeRecordMerge");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, cancel);
				return;
			}

			NetOffice.PublisherApi._Document newDoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.PublisherApi._Document;
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("MailMergeBeforeRecordMerge", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void MailMergeDataSourceLoad([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MailMergeDataSourceLoad");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.PublisherApi._Document newDoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.PublisherApi._Document;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			_eventBinding.RaiseCustomEvent("MailMergeDataSourceLoad", ref paramsArray);
		}

		public void MailMergeWizardSendToCustom([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MailMergeWizardSendToCustom");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.PublisherApi._Document newDoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.PublisherApi._Document;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			_eventBinding.RaiseCustomEvent("MailMergeWizardSendToCustom", ref paramsArray);
		}

		public void MailMergeWizardStateChange([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object fromState)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MailMergeWizardStateChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, fromState);
				return;
			}

			NetOffice.PublisherApi._Document newDoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.PublisherApi._Document;
			Int32 newFromState = Convert.ToInt32(fromState);
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray[1] = newFromState;
			_eventBinding.RaiseCustomEvent("MailMergeWizardStateChange", ref paramsArray);
		}

		public void MailMergeDataSourceValidate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object handled)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MailMergeDataSourceValidate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, handled);
				return;
			}

			NetOffice.PublisherApi._Document newDoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.PublisherApi._Document;
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(handled, 1);
			_eventBinding.RaiseCustomEvent("MailMergeDataSourceValidate", ref paramsArray);

			handled = (bool)paramsArray[1];
		}

		public void MailMergeInsertBarcode([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object okToInsert)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MailMergeInsertBarcode");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, okToInsert);
				return;
			}

			NetOffice.PublisherApi._Document newDoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.PublisherApi._Document;
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(okToInsert, 1);
			_eventBinding.RaiseCustomEvent("MailMergeInsertBarcode", ref paramsArray);

			okToInsert = (bool)paramsArray[1];
		}

		public void MailMergeRecipientListClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MailMergeRecipientListClose");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.PublisherApi._Document newDoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.PublisherApi._Document;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			_eventBinding.RaiseCustomEvent("MailMergeRecipientListClose", ref paramsArray);
		}

		public void MailMergeGenerateBarcode([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object bstrString)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MailMergeGenerateBarcode");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, bstrString);
				return;
			}

			NetOffice.PublisherApi._Document newDoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.PublisherApi._Document;
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(bstrString, 1);
			_eventBinding.RaiseCustomEvent("MailMergeGenerateBarcode", ref paramsArray);

			bstrString = (string)paramsArray[1];
		}

		public void MailMergeWizardFollowUpCustom([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MailMergeWizardFollowUpCustom");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.PublisherApi._Document newDoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.PublisherApi._Document;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			_eventBinding.RaiseCustomEvent("MailMergeWizardFollowUpCustom", ref paramsArray);
		}

		public void BeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforePrint");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, cancel);
				return;
			}

			NetOffice.PublisherApi._Document newDoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.PublisherApi._Document;
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("BeforePrint", ref paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void AfterPrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterPrint");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.PublisherApi._Document newDoc = Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.PublisherApi._Document;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			_eventBinding.RaiseCustomEvent("AfterPrint", ref paramsArray);
		}

		public void ShowCatalogUI()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ShowCatalogUI");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("ShowCatalogUI", ref paramsArray);
		}

		public void HideCatalogUI()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("HideCatalogUI");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("HideCatalogUI", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}