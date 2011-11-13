using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace NetOffice.WordApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibraryAttribute("Word", 11,12,14)]
	[ComImport, Guid("00020A01-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ApplicationEvents4
	{
		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Startup();

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void Quit();

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void DocumentChange();

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void DocumentOpen([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void DocumentBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void DocumentBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void DocumentBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object saveAsUI, [In] [Out] ref object cancel);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void NewDocument([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
		void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
		void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void WindowBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(14)]
		void WindowBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(15)]
		void EPostagePropertyDialog([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16)]
		void EPostageInsert([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(17)]
		void MailMergeAfterMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object docResult);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(18)]
		void MailMergeAfterRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(19)]
		void MailMergeBeforeMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object startRecord, [In] object endRecord, [In] [Out] ref object cancel);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(20)]
		void MailMergeBeforeRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(21)]
		void MailMergeDataSourceLoad([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(22)]
		void MailMergeDataSourceValidate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object handled);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(23)]
		void MailMergeWizardSendToCustom([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(24)]
		void MailMergeWizardStateChange([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object fromState, [In] [Out] ref object toState, [In] [Out] ref object handled);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(25)]
		void WindowSize([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(26)]
		void XMLSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In, MarshalAs(UnmanagedType.IDispatch)] object oldXMLNode, [In, MarshalAs(UnmanagedType.IDispatch)] object newXMLNode, [In] [Out] ref object reason);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(27)]
		void XMLValidationError([In, MarshalAs(UnmanagedType.IDispatch)] object xMLNode);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(28)]
		void DocumentSync([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object syncEventType);

		[SupportByLibraryAttribute("Word", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(29)]
		void EPostageInsertEx([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object cpDeliveryAddrStart, [In] object cpDeliveryAddrEnd, [In] object cpReturnAddrStart, [In] object cpReturnAddrEnd, [In] object xaWidth, [In] object yaHeight, [In] object bstrPrinterName, [In] object bstrPaperFeed, [In] object fPrint, [In] [Out] ref object fCancel);

		[SupportByLibraryAttribute("Word", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(30)]
		void MailMergeDataSourceValidate2([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object handled);

		[SupportByLibraryAttribute("Word", 14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(31)]
		void ProtectedViewWindowOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow);

		[SupportByLibraryAttribute("Word", 14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(32)]
		void ProtectedViewWindowBeforeEdit([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow, [In] [Out] ref object cancel);

		[SupportByLibraryAttribute("Word", 14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(33)]
		void ProtectedViewWindowBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow, [In] object closeReason, [In] [Out] ref object cancel);

		[SupportByLibraryAttribute("Word", 14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(34)]
		void ProtectedViewWindowSize([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow);

		[SupportByLibraryAttribute("Word", 14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(35)]
		void ProtectedViewWindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow);

		[SupportByLibraryAttribute("Word", 14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(36)]
		void ProtectedViewWindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ApplicationEvents4_SinkHelper : SinkHelper, ApplicationEvents4
	{
		#region Static
		
		public static readonly string Id = "00020A01-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public ApplicationEvents4_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region ApplicationEvents4 Members
		
		public void Startup()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Startup");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
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
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void DocumentChange()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DocumentChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void DocumentOpen([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DocumentOpen");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void DocumentBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DocumentBeforeClose");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, cancel);
				return;
			}

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(cancel, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void DocumentBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DocumentBeforePrint");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, cancel);
				return;
			}

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(cancel, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void DocumentBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object saveAsUI, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DocumentBeforeSave");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, saveAsUI, cancel);
				return;
			}

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			object[] paramsArray = new object[3];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(saveAsUI, 1);
			paramsArray.SetValue(cancel, 2);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			saveAsUI = (bool)paramsArray[1];
			cancel = (bool)paramsArray[2];
		}

		public void NewDocument([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("NewDocument");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowActivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, wn);
				return;
			}

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			NetOffice.WordApi.Window newWn = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wn) as NetOffice.WordApi.Window;
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray[1] = newWn;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowDeactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, wn);
				return;
			}

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			NetOffice.WordApi.Window newWn = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wn) as NetOffice.WordApi.Window;
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray[1] = newWn;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowSelectionChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sel);
				return;
			}

			NetOffice.WordApi.Selection newSel = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sel) as NetOffice.WordApi.Selection;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSel;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WindowBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowBeforeRightClick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sel, cancel);
				return;
			}

			NetOffice.WordApi.Selection newSel = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sel) as NetOffice.WordApi.Selection;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSel;
			paramsArray.SetValue(cancel, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void WindowBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowBeforeDoubleClick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sel, cancel);
				return;
			}

			NetOffice.WordApi.Selection newSel = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sel) as NetOffice.WordApi.Selection;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSel;
			paramsArray.SetValue(cancel, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void EPostagePropertyDialog([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("EPostagePropertyDialog");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void EPostageInsert([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("EPostageInsert");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void MailMergeAfterMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object docResult)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MailMergeAfterMerge");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, docResult);
				return;
			}

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			NetOffice.WordApi.Document newDocResult = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, docResult) as NetOffice.WordApi.Document;
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray[1] = newDocResult;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void MailMergeAfterRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MailMergeAfterRecordMerge");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void MailMergeBeforeMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object startRecord, [In] object endRecord, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MailMergeBeforeMerge");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, startRecord, endRecord, cancel);
				return;
			}

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			Int32 newStartRecord = (Int32)startRecord;
			Int32 newEndRecord = (Int32)endRecord;
			object[] paramsArray = new object[4];
			paramsArray[0] = newDoc;
			paramsArray[1] = newStartRecord;
			paramsArray[2] = newEndRecord;
			paramsArray.SetValue(cancel, 3);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

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

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(cancel, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

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

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void MailMergeDataSourceValidate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object handled)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MailMergeDataSourceValidate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, handled);
				return;
			}

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(handled, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			handled = (bool)paramsArray[1];
		}

		public void MailMergeWizardSendToCustom([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MailMergeWizardSendToCustom");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc);
				return;
			}

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			object[] paramsArray = new object[1];
			paramsArray[0] = newDoc;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void MailMergeWizardStateChange([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object fromState, [In] [Out] ref object toState, [In] [Out] ref object handled)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MailMergeWizardStateChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, fromState, toState, handled);
				return;
			}

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			object[] paramsArray = new object[4];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(fromState, 1);
			paramsArray.SetValue(toState, 2);
			paramsArray.SetValue(handled, 3);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			fromState = (Int32)paramsArray[1];
			toState = (Int32)paramsArray[2];
			handled = (bool)paramsArray[3];
		}

		public void WindowSize([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowSize");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, wn);
				return;
			}

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			NetOffice.WordApi.Window newWn = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wn) as NetOffice.WordApi.Window;
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray[1] = newWn;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void XMLSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In, MarshalAs(UnmanagedType.IDispatch)] object oldXMLNode, [In, MarshalAs(UnmanagedType.IDispatch)] object newXMLNode, [In] [Out] ref object reason)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("XMLSelectionChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sel, oldXMLNode, newXMLNode, reason);
				return;
			}

			NetOffice.WordApi.Selection newSel = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sel) as NetOffice.WordApi.Selection;
			NetOffice.WordApi.XMLNode newOldXMLNode = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, oldXMLNode) as NetOffice.WordApi.XMLNode;
			NetOffice.WordApi.XMLNode newNewXMLNode = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, newXMLNode) as NetOffice.WordApi.XMLNode;
			object[] paramsArray = new object[4];
			paramsArray[0] = newSel;
			paramsArray[1] = newOldXMLNode;
			paramsArray[2] = newNewXMLNode;
			paramsArray.SetValue(reason, 3);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			reason = (Int32)paramsArray[3];
		}

		public void XMLValidationError([In, MarshalAs(UnmanagedType.IDispatch)] object xMLNode)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("XMLValidationError");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(xMLNode);
				return;
			}

			NetOffice.WordApi.XMLNode newXMLNode = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, xMLNode) as NetOffice.WordApi.XMLNode;
			object[] paramsArray = new object[1];
			paramsArray[0] = newXMLNode;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void DocumentSync([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object syncEventType)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DocumentSync");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, syncEventType);
				return;
			}

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			NetOffice.OfficeApi.Enums.MsoSyncEventType newSyncEventType = (NetOffice.OfficeApi.Enums.MsoSyncEventType)syncEventType;
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray[1] = newSyncEventType;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void EPostageInsertEx([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object cpDeliveryAddrStart, [In] object cpDeliveryAddrEnd, [In] object cpReturnAddrStart, [In] object cpReturnAddrEnd, [In] object xaWidth, [In] object yaHeight, [In] object bstrPrinterName, [In] object bstrPaperFeed, [In] object fPrint, [In] [Out] ref object fCancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("EPostageInsertEx");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, cpDeliveryAddrStart, cpDeliveryAddrEnd, cpReturnAddrStart, cpReturnAddrEnd, xaWidth, yaHeight, bstrPrinterName, bstrPaperFeed, fPrint, fCancel);
				return;
			}

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			Int32 newcpDeliveryAddrStart = (Int32)cpDeliveryAddrStart;
			Int32 newcpDeliveryAddrEnd = (Int32)cpDeliveryAddrEnd;
			Int32 newcpReturnAddrStart = (Int32)cpReturnAddrStart;
			Int32 newcpReturnAddrEnd = (Int32)cpReturnAddrEnd;
			Int32 newxaWidth = (Int32)xaWidth;
			Int32 newyaHeight = (Int32)yaHeight;
			string newbstrPrinterName = (string)bstrPrinterName;
			string newbstrPaperFeed = (string)bstrPaperFeed;
			bool newfPrint = (bool)fPrint;
			object[] paramsArray = new object[11];
			paramsArray[0] = newDoc;
			paramsArray[1] = newcpDeliveryAddrStart;
			paramsArray[2] = newcpDeliveryAddrEnd;
			paramsArray[3] = newcpReturnAddrStart;
			paramsArray[4] = newcpReturnAddrEnd;
			paramsArray[5] = newxaWidth;
			paramsArray[6] = newyaHeight;
			paramsArray[7] = newbstrPrinterName;
			paramsArray[8] = newbstrPaperFeed;
			paramsArray[9] = newfPrint;
			paramsArray.SetValue(fCancel, 10);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			fCancel = (bool)paramsArray[10];
		}

		public void MailMergeDataSourceValidate2([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object handled)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MailMergeDataSourceValidate2");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(doc, handled);
				return;
			}

			NetOffice.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as NetOffice.WordApi.Document;
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray.SetValue(handled, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			handled = (bool)paramsArray[1];
		}

		public void ProtectedViewWindowOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProtectedViewWindowOpen");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pvWindow);
				return;
			}

			NetOffice.WordApi.ProtectedViewWindow newPvWindow = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pvWindow) as NetOffice.WordApi.ProtectedViewWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPvWindow;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void ProtectedViewWindowBeforeEdit([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProtectedViewWindowBeforeEdit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pvWindow, cancel);
				return;
			}

			NetOffice.WordApi.ProtectedViewWindow newPvWindow = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pvWindow) as NetOffice.WordApi.ProtectedViewWindow;
			object[] paramsArray = new object[2];
			paramsArray[0] = newPvWindow;
			paramsArray.SetValue(cancel, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void ProtectedViewWindowBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow, [In] object closeReason, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProtectedViewWindowBeforeClose");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pvWindow, closeReason, cancel);
				return;
			}

			NetOffice.WordApi.ProtectedViewWindow newPvWindow = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pvWindow) as NetOffice.WordApi.ProtectedViewWindow;
			Int32 newCloseReason = (Int32)closeReason;
			object[] paramsArray = new object[3];
			paramsArray[0] = newPvWindow;
			paramsArray[1] = newCloseReason;
			paramsArray.SetValue(cancel, 2);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[2];
		}

		public void ProtectedViewWindowSize([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProtectedViewWindowSize");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pvWindow);
				return;
			}

			NetOffice.WordApi.ProtectedViewWindow newPvWindow = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pvWindow) as NetOffice.WordApi.ProtectedViewWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPvWindow;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void ProtectedViewWindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProtectedViewWindowActivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pvWindow);
				return;
			}

			NetOffice.WordApi.ProtectedViewWindow newPvWindow = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pvWindow) as NetOffice.WordApi.ProtectedViewWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPvWindow;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void ProtectedViewWindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvWindow)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProtectedViewWindowDeactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pvWindow);
				return;
			}

			NetOffice.WordApi.ProtectedViewWindow newPvWindow = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pvWindow) as NetOffice.WordApi.ProtectedViewWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPvWindow;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}