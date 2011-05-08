using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.WordApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibrary("WD10","WD11","WD12","WD14")]
	[ComImport, Guid("00020A00-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ApplicationEvents3
	{
		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Startup();

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void Quit();

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void DocumentChange();

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void DocumentOpen([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void DocumentBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void DocumentBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void DocumentBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object saveAsUI, [In] [Out] ref object cancel);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void NewDocument([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
		void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
		void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void WindowBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(14)]
		void WindowBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(15)]
		void EPostagePropertyDialog([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(16)]
		void EPostageInsert([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(17)]
		void MailMergeAfterMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object docResult);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(18)]
		void MailMergeAfterRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(19)]
		void MailMergeBeforeMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] object startRecord, [In] object endRecord, [In] [Out] ref object cancel);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(20)]
		void MailMergeBeforeRecordMerge([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(21)]
		void MailMergeDataSourceLoad([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(22)]
		void MailMergeDataSourceValidate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object handled);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(23)]
		void MailMergeWizardSendToCustom([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(24)]
		void MailMergeWizardStateChange([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object fromState, [In] [Out] ref object toState, [In] [Out] ref object handled);

		[SupportByLibrary("WD10","WD11","WD12","WD14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(25)]
		void WindowSize([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ApplicationEvents3_SinkHelper : SinkHelper, ApplicationEvents3
	{
		#region Static
		
		public static readonly string Id = "00020A00-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public ApplicationEvents3_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region ApplicationEvents3 Members
		
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

			LateBindingApi.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as LateBindingApi.WordApi.Document;
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

			LateBindingApi.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as LateBindingApi.WordApi.Document;
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

			LateBindingApi.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as LateBindingApi.WordApi.Document;
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

			LateBindingApi.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as LateBindingApi.WordApi.Document;
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

			LateBindingApi.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as LateBindingApi.WordApi.Document;
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

			LateBindingApi.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as LateBindingApi.WordApi.Document;
			LateBindingApi.WordApi.Window newWn = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wn) as LateBindingApi.WordApi.Window;
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

			LateBindingApi.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as LateBindingApi.WordApi.Document;
			LateBindingApi.WordApi.Window newWn = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wn) as LateBindingApi.WordApi.Window;
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

			LateBindingApi.WordApi.Selection newSel = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sel) as LateBindingApi.WordApi.Selection;
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

			LateBindingApi.WordApi.Selection newSel = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sel) as LateBindingApi.WordApi.Selection;
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

			LateBindingApi.WordApi.Selection newSel = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sel) as LateBindingApi.WordApi.Selection;
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

			LateBindingApi.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as LateBindingApi.WordApi.Document;
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

			LateBindingApi.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as LateBindingApi.WordApi.Document;
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

			LateBindingApi.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as LateBindingApi.WordApi.Document;
			LateBindingApi.WordApi.Document newDocResult = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, docResult) as LateBindingApi.WordApi.Document;
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

			LateBindingApi.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as LateBindingApi.WordApi.Document;
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

			LateBindingApi.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as LateBindingApi.WordApi.Document;
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

			LateBindingApi.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as LateBindingApi.WordApi.Document;
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

			LateBindingApi.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as LateBindingApi.WordApi.Document;
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

			LateBindingApi.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as LateBindingApi.WordApi.Document;
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

			LateBindingApi.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as LateBindingApi.WordApi.Document;
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

			LateBindingApi.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as LateBindingApi.WordApi.Document;
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

			LateBindingApi.WordApi.Document newDoc = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, doc) as LateBindingApi.WordApi.Document;
			LateBindingApi.WordApi.Window newWn = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wn) as LateBindingApi.WordApi.Window;
			object[] paramsArray = new object[2];
			paramsArray[0] = newDoc;
			paramsArray[1] = newWn;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}