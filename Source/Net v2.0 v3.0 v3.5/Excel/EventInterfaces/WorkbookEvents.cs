using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.ExcelApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
	[ComImport, Guid("00024412-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface WorkbookEvents
	{
		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(682)]
		void Open();

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(304)]
		void Activate();

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1530)]
		void Deactivate();

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1546)]
		void BeforeClose([In] [Out] ref object cancel);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1547)]
		void BeforeSave([In] object saveAsUI, [In] [Out] ref object cancel);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1549)]
		void BeforePrint([In] [Out] ref object cancel);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1550)]
		void NewSheet([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1552)]
		void AddinInstall();

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1553)]
		void AddinUninstall();

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1554)]
		void WindowResize([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1556)]
		void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1557)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1558)]
		void SheetSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1559)]
		void SheetBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1560)]
		void SheetBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1561)]
		void SheetActivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1562)]
		void SheetDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1563)]
		void SheetCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1564)]
		void SheetChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1854)]
		void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2157)]
		void SheetPivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2158)]
		void PivotTableCloseConnection([In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2159)]
		void PivotTableOpenConnection([In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibrary("XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2266)]
		void Sync([In] object syncEventType);

		[SupportByLibrary("XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2283)]
		void BeforeXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object isRefresh, [In] [Out] ref object cancel);

		[SupportByLibrary("XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2285)]
		void AfterXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object isRefresh, [In] object result);

		[SupportByLibrary("XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2287)]
		void BeforeXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] [Out] ref object cancel);

		[SupportByLibrary("XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2288)]
		void AfterXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object result);

		[SupportByLibrary("XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2610)]
		void RowsetComplete([In] object description, [In] object sheet, [In] object success);

		[SupportByLibrary("XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2895)]
		void SheetPivotTableAfterValueChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In, MarshalAs(UnmanagedType.IDispatch)] object targetRange);

		[SupportByLibrary("XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2896)]
		void SheetPivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel);

		[SupportByLibrary("XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2897)]
		void SheetPivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel);

		[SupportByLibrary("XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2898)]
		void SheetPivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd);

		[SupportByLibrary("XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2899)]
		void SheetPivotTableChangeSync([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibrary("XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2900)]
		void AfterSave([In] object success);

		[SupportByLibrary("XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2901)]
		void NewChart([In, MarshalAs(UnmanagedType.IDispatch)] object ch);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class WorkbookEvents_SinkHelper : SinkHelper, WorkbookEvents
	{
		#region Static
		
		public static readonly string Id = "00024412-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public WorkbookEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region WorkbookEvents Members
		
		public void Open()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Open");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void Activate()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Activate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void Deactivate()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Deactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void BeforeClose([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeClose");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[0];
		}

		public void BeforeSave([In] object saveAsUI, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeSave");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(saveAsUI, cancel);
				return;
			}

			bool newSaveAsUI = (bool)saveAsUI;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSaveAsUI;
			paramsArray.SetValue(cancel, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void BeforePrint([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforePrint");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[0];
		}

		public void NewSheet([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("NewSheet");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh);
				return;
			}

			object newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSh;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void AddinInstall()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AddinInstall");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void AddinUninstall()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AddinUninstall");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WindowResize([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowResize");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wn);
				return;
			}

			LateBindingApi.ExcelApi.Window newWn = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wn) as LateBindingApi.ExcelApi.Window;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowActivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wn);
				return;
			}

			LateBindingApi.ExcelApi.Window newWn = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wn) as LateBindingApi.ExcelApi.Window;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowDeactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wn);
				return;
			}

			LateBindingApi.ExcelApi.Window newWn = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wn) as LateBindingApi.ExcelApi.Window;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void SheetSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetSelectionChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, target);
				return;
			}

			object newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			LateBindingApi.ExcelApi.Range newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as LateBindingApi.ExcelApi.Range;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSh;
			paramsArray[1] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void SheetBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetBeforeDoubleClick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, target, cancel);
				return;
			}

			object newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			LateBindingApi.ExcelApi.Range newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as LateBindingApi.ExcelApi.Range;
			object[] paramsArray = new object[3];
			paramsArray[0] = newSh;
			paramsArray[1] = newTarget;
			paramsArray.SetValue(cancel, 2);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[2];
		}

		public void SheetBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetBeforeRightClick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, target, cancel);
				return;
			}

			object newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			LateBindingApi.ExcelApi.Range newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as LateBindingApi.ExcelApi.Range;
			object[] paramsArray = new object[3];
			paramsArray[0] = newSh;
			paramsArray[1] = newTarget;
			paramsArray.SetValue(cancel, 2);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[2];
		}

		public void SheetActivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetActivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh);
				return;
			}

			object newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSh;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void SheetDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetDeactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh);
				return;
			}

			object newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSh;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void SheetCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetCalculate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh);
				return;
			}

			object newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSh;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void SheetChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, target);
				return;
			}

			object newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			LateBindingApi.ExcelApi.Range newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as LateBindingApi.ExcelApi.Range;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSh;
			paramsArray[1] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetFollowHyperlink");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, target);
				return;
			}

			object newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			LateBindingApi.ExcelApi.Hyperlink newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as LateBindingApi.ExcelApi.Hyperlink;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSh;
			paramsArray[1] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void SheetPivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetPivotTableUpdate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, target);
				return;
			}

			object newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			LateBindingApi.ExcelApi.PivotTable newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as LateBindingApi.ExcelApi.PivotTable;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSh;
			paramsArray[1] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void PivotTableCloseConnection([In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PivotTableCloseConnection");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(target);
				return;
			}

			LateBindingApi.ExcelApi.PivotTable newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as LateBindingApi.ExcelApi.PivotTable;
			object[] paramsArray = new object[1];
			paramsArray[0] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void PivotTableOpenConnection([In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PivotTableOpenConnection");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(target);
				return;
			}

			LateBindingApi.ExcelApi.PivotTable newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as LateBindingApi.ExcelApi.PivotTable;
			object[] paramsArray = new object[1];
			paramsArray[0] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void Sync([In] object syncEventType)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Sync");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(syncEventType);
				return;
			}

			LateBindingApi.OfficeApi.Enums.MsoSyncEventType newSyncEventType = (LateBindingApi.OfficeApi.Enums.MsoSyncEventType)syncEventType;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSyncEventType;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void BeforeXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object isRefresh, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeXmlImport");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(map, url, isRefresh, cancel);
				return;
			}

			LateBindingApi.ExcelApi.XmlMap newMap = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, map) as LateBindingApi.ExcelApi.XmlMap;
			string newUrl = (string)url;
			bool newIsRefresh = (bool)isRefresh;
			object[] paramsArray = new object[4];
			paramsArray[0] = newMap;
			paramsArray[1] = newUrl;
			paramsArray[2] = newIsRefresh;
			paramsArray.SetValue(cancel, 3);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[3];
		}

		public void AfterXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object isRefresh, [In] object result)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterXmlImport");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(map, isRefresh, result);
				return;
			}

			LateBindingApi.ExcelApi.XmlMap newMap = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, map) as LateBindingApi.ExcelApi.XmlMap;
			bool newIsRefresh = (bool)isRefresh;
			LateBindingApi.ExcelApi.Enums.XlXmlImportResult newResult = (LateBindingApi.ExcelApi.Enums.XlXmlImportResult)result;
			object[] paramsArray = new object[3];
			paramsArray[0] = newMap;
			paramsArray[1] = newIsRefresh;
			paramsArray[2] = newResult;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void BeforeXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeXmlExport");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(map, url, cancel);
				return;
			}

			LateBindingApi.ExcelApi.XmlMap newMap = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, map) as LateBindingApi.ExcelApi.XmlMap;
			string newUrl = (string)url;
			object[] paramsArray = new object[3];
			paramsArray[0] = newMap;
			paramsArray[1] = newUrl;
			paramsArray.SetValue(cancel, 2);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[2];
		}

		public void AfterXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object result)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterXmlExport");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(map, url, result);
				return;
			}

			LateBindingApi.ExcelApi.XmlMap newMap = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, map) as LateBindingApi.ExcelApi.XmlMap;
			string newUrl = (string)url;
			LateBindingApi.ExcelApi.Enums.XlXmlExportResult newResult = (LateBindingApi.ExcelApi.Enums.XlXmlExportResult)result;
			object[] paramsArray = new object[3];
			paramsArray[0] = newMap;
			paramsArray[1] = newUrl;
			paramsArray[2] = newResult;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void RowsetComplete([In] object description, [In] object sheet, [In] object success)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("RowsetComplete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(description, sheet, success);
				return;
			}

			string newDescription = (string)description;
			string newSheet = (string)sheet;
			bool newSuccess = (bool)success;
			object[] paramsArray = new object[3];
			paramsArray[0] = newDescription;
			paramsArray[1] = newSheet;
			paramsArray[2] = newSuccess;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void SheetPivotTableAfterValueChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In, MarshalAs(UnmanagedType.IDispatch)] object targetRange)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetPivotTableAfterValueChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, targetPivotTable, targetRange);
				return;
			}

			object newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			LateBindingApi.ExcelApi.PivotTable newTargetPivotTable = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as LateBindingApi.ExcelApi.PivotTable;
			LateBindingApi.ExcelApi.Range newTargetRange = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, targetRange) as LateBindingApi.ExcelApi.Range;
			object[] paramsArray = new object[3];
			paramsArray[0] = newSh;
			paramsArray[1] = newTargetPivotTable;
			paramsArray[2] = newTargetRange;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void SheetPivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetPivotTableBeforeAllocateChanges");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, targetPivotTable, valueChangeStart, valueChangeEnd, cancel);
				return;
			}

			object newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			LateBindingApi.ExcelApi.PivotTable newTargetPivotTable = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as LateBindingApi.ExcelApi.PivotTable;
			Int32 newValueChangeStart = (Int32)valueChangeStart;
			Int32 newValueChangeEnd = (Int32)valueChangeEnd;
			object[] paramsArray = new object[5];
			paramsArray[0] = newSh;
			paramsArray[1] = newTargetPivotTable;
			paramsArray[2] = newValueChangeStart;
			paramsArray[3] = newValueChangeEnd;
			paramsArray.SetValue(cancel, 4);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[4];
		}

		public void SheetPivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetPivotTableBeforeCommitChanges");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, targetPivotTable, valueChangeStart, valueChangeEnd, cancel);
				return;
			}

			object newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			LateBindingApi.ExcelApi.PivotTable newTargetPivotTable = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as LateBindingApi.ExcelApi.PivotTable;
			Int32 newValueChangeStart = (Int32)valueChangeStart;
			Int32 newValueChangeEnd = (Int32)valueChangeEnd;
			object[] paramsArray = new object[5];
			paramsArray[0] = newSh;
			paramsArray[1] = newTargetPivotTable;
			paramsArray[2] = newValueChangeStart;
			paramsArray[3] = newValueChangeEnd;
			paramsArray.SetValue(cancel, 4);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[4];
		}

		public void SheetPivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetPivotTableBeforeDiscardChanges");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, targetPivotTable, valueChangeStart, valueChangeEnd);
				return;
			}

			object newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			LateBindingApi.ExcelApi.PivotTable newTargetPivotTable = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as LateBindingApi.ExcelApi.PivotTable;
			Int32 newValueChangeStart = (Int32)valueChangeStart;
			Int32 newValueChangeEnd = (Int32)valueChangeEnd;
			object[] paramsArray = new object[4];
			paramsArray[0] = newSh;
			paramsArray[1] = newTargetPivotTable;
			paramsArray[2] = newValueChangeStart;
			paramsArray[3] = newValueChangeEnd;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void SheetPivotTableChangeSync([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetPivotTableChangeSync");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, target);
				return;
			}

			object newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			LateBindingApi.ExcelApi.PivotTable newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as LateBindingApi.ExcelApi.PivotTable;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSh;
			paramsArray[1] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void AfterSave([In] object success)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterSave");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(success);
				return;
			}

			bool newSuccess = (bool)success;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSuccess;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void NewChart([In, MarshalAs(UnmanagedType.IDispatch)] object ch)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("NewChart");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(ch);
				return;
			}

			LateBindingApi.ExcelApi.Chart newCh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, ch) as LateBindingApi.ExcelApi.Chart;
			object[] paramsArray = new object[1];
			paramsArray[0] = newCh;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}