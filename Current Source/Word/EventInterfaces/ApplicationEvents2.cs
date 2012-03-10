using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace NetOffice.WordApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibraryAttribute("Word", 9,10,11,12,14)]
	[ComImport, Guid("000209FE-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ApplicationEvents2
	{
		[SupportByLibraryAttribute("Word", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Startup();

		[SupportByLibraryAttribute("Word", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void Quit();

		[SupportByLibraryAttribute("Word", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void DocumentChange();

		[SupportByLibraryAttribute("Word", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void DocumentOpen([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByLibraryAttribute("Word", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
		void DocumentBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		[SupportByLibraryAttribute("Word", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(7)]
		void DocumentBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel);

		[SupportByLibraryAttribute("Word", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(8)]
		void DocumentBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object saveAsUI, [In] [Out] ref object cancel);

		[SupportByLibraryAttribute("Word", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void NewDocument([In, MarshalAs(UnmanagedType.IDispatch)] object doc);

		[SupportByLibraryAttribute("Word", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
		void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByLibraryAttribute("Word", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByLibraryAttribute("Word", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
		void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel);

		[SupportByLibraryAttribute("Word", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
		void WindowBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);

		[SupportByLibraryAttribute("Word", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(14)]
		void WindowBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ApplicationEvents2_SinkHelper : SinkHelper, ApplicationEvents2
	{
		#region Static
		
		public static readonly string Id = "000209FE-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public ApplicationEvents2_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region ApplicationEvents2 Members
		
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

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}