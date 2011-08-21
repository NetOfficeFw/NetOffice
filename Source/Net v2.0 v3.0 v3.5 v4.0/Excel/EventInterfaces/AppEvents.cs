using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace NetOffice.ExcelApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibrary("Excel", 9,10,11,12,14)]
	[ComImport, Guid("00024413-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface AppEvents
	{
		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1565)]
		void NewWorkbook([In, MarshalAs(UnmanagedType.IDispatch)] object wb);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1558)]
		void SheetSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1559)]
		void SheetBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1560)]
		void SheetBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1561)]
		void SheetActivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1562)]
		void SheetDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1563)]
		void SheetCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1564)]
		void SheetChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1567)]
		void WorkbookOpen([In, MarshalAs(UnmanagedType.IDispatch)] object wb);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1568)]
		void WorkbookActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1569)]
		void WorkbookDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1570)]
		void WorkbookBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] [Out] ref object cancel);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1571)]
		void WorkbookBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object saveAsUI, [In] [Out] ref object cancel);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1572)]
		void WorkbookBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] [Out] ref object cancel);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1573)]
		void WorkbookNewSheet([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1574)]
		void WorkbookAddinInstall([In, MarshalAs(UnmanagedType.IDispatch)] object wb);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1575)]
		void WorkbookAddinUninstall([In, MarshalAs(UnmanagedType.IDispatch)] object wb);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1554)]
		void WindowResize([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1556)]
		void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1557)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByLibrary("Excel", 9,10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1854)]
		void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibrary("Excel", 10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2157)]
		void SheetPivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibrary("Excel", 10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2160)]
		void WorkbookPivotTableCloseConnection([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibrary("Excel", 10,11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2161)]
		void WorkbookPivotTableOpenConnection([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibrary("Excel", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2289)]
		void WorkbookSync([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object syncEventType);

		[SupportByLibrary("Excel", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2290)]
		void WorkbookBeforeXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object isRefresh, [In] [Out] ref object cancel);

		[SupportByLibrary("Excel", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2291)]
		void WorkbookAfterXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object isRefresh, [In] object result);

		[SupportByLibrary("Excel", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2292)]
		void WorkbookBeforeXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] [Out] ref object cancel);

		[SupportByLibrary("Excel", 11,12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2293)]
		void WorkbookAfterXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object result);

		[SupportByLibrary("Excel", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2611)]
		void WorkbookRowsetComplete([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object description, [In] object sheet, [In] object success);

		[SupportByLibrary("Excel", 12,14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2612)]
		void AfterCalculate();

		[SupportByLibrary("Excel", 14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2895)]
		void SheetPivotTableAfterValueChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In, MarshalAs(UnmanagedType.IDispatch)] object targetRange);

		[SupportByLibrary("Excel", 14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2896)]
		void SheetPivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel);

		[SupportByLibrary("Excel", 14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2897)]
		void SheetPivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel);

		[SupportByLibrary("Excel", 14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2898)]
		void SheetPivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd);

		[SupportByLibrary("Excel", 14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2903)]
		void ProtectedViewWindowOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pvw);

		[SupportByLibrary("Excel", 14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2905)]
		void ProtectedViewWindowBeforeEdit([In, MarshalAs(UnmanagedType.IDispatch)] object pvw, [In] [Out] ref object cancel);

		[SupportByLibrary("Excel", 14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2906)]
		void ProtectedViewWindowBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pvw, [In] object reason, [In] [Out] ref object cancel);

		[SupportByLibrary("Excel", 14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2908)]
		void ProtectedViewWindowResize([In, MarshalAs(UnmanagedType.IDispatch)] object pvw);

		[SupportByLibrary("Excel", 14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2909)]
		void ProtectedViewWindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvw);

		[SupportByLibrary("Excel", 14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2910)]
		void ProtectedViewWindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvw);

		[SupportByLibrary("Excel", 14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2911)]
		void WorkbookAfterSave([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object success);

		[SupportByLibrary("Excel", 14)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2912)]
		void WorkbookNewChart([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object ch);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class AppEvents_SinkHelper : SinkHelper, AppEvents
	{
		#region Static
		
		public static readonly string Id = "00024413-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public AppEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region AppEvents Members
		
		public void NewWorkbook([In, MarshalAs(UnmanagedType.IDispatch)] object wb)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("NewWorkbook");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWb;
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
			NetOffice.ExcelApi.Range newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Range;
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
			NetOffice.ExcelApi.Range newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Range;
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
			NetOffice.ExcelApi.Range newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Range;
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
			NetOffice.ExcelApi.Range newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Range;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSh;
			paramsArray[1] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WorkbookOpen([In, MarshalAs(UnmanagedType.IDispatch)] object wb)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkbookOpen");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWb;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WorkbookActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkbookActivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWb;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WorkbookDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkbookDeactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWb;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WorkbookBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkbookBeforeClose");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb, cancel);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			object[] paramsArray = new object[2];
			paramsArray[0] = newWb;
			paramsArray.SetValue(cancel, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void WorkbookBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object saveAsUI, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkbookBeforeSave");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb, saveAsUI, cancel);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			bool newSaveAsUI = (bool)saveAsUI;
			object[] paramsArray = new object[3];
			paramsArray[0] = newWb;
			paramsArray[1] = newSaveAsUI;
			paramsArray.SetValue(cancel, 2);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[2];
		}

		public void WorkbookBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkbookBeforePrint");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb, cancel);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			object[] paramsArray = new object[2];
			paramsArray[0] = newWb;
			paramsArray.SetValue(cancel, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void WorkbookNewSheet([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object sh)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkbookNewSheet");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb, sh);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			object newSh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newWb;
			paramsArray[1] = newSh;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WorkbookAddinInstall([In, MarshalAs(UnmanagedType.IDispatch)] object wb)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkbookAddinInstall");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWb;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WorkbookAddinUninstall([In, MarshalAs(UnmanagedType.IDispatch)] object wb)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkbookAddinUninstall");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWb;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WindowResize([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowResize");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb, wn);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			NetOffice.ExcelApi.Window newWn = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wn) as NetOffice.ExcelApi.Window;
			object[] paramsArray = new object[2];
			paramsArray[0] = newWb;
			paramsArray[1] = newWn;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowActivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb, wn);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			NetOffice.ExcelApi.Window newWn = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wn) as NetOffice.ExcelApi.Window;
			object[] paramsArray = new object[2];
			paramsArray[0] = newWb;
			paramsArray[1] = newWn;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowDeactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb, wn);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			NetOffice.ExcelApi.Window newWn = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wn) as NetOffice.ExcelApi.Window;
			object[] paramsArray = new object[2];
			paramsArray[0] = newWb;
			paramsArray[1] = newWn;
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
			NetOffice.ExcelApi.Hyperlink newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Hyperlink;
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
			NetOffice.ExcelApi.PivotTable newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.PivotTable;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSh;
			paramsArray[1] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WorkbookPivotTableCloseConnection([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkbookPivotTableCloseConnection");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb, target);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			NetOffice.ExcelApi.PivotTable newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.PivotTable;
			object[] paramsArray = new object[2];
			paramsArray[0] = newWb;
			paramsArray[1] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WorkbookPivotTableOpenConnection([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkbookPivotTableOpenConnection");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb, target);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			NetOffice.ExcelApi.PivotTable newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.PivotTable;
			object[] paramsArray = new object[2];
			paramsArray[0] = newWb;
			paramsArray[1] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WorkbookSync([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object syncEventType)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkbookSync");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb, syncEventType);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			NetOffice.OfficeApi.Enums.MsoSyncEventType newSyncEventType = (NetOffice.OfficeApi.Enums.MsoSyncEventType)syncEventType;
			object[] paramsArray = new object[2];
			paramsArray[0] = newWb;
			paramsArray[1] = newSyncEventType;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WorkbookBeforeXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object isRefresh, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkbookBeforeXmlImport");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb, map, url, isRefresh, cancel);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			NetOffice.ExcelApi.XmlMap newMap = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, map) as NetOffice.ExcelApi.XmlMap;
			string newUrl = (string)url;
			bool newIsRefresh = (bool)isRefresh;
			object[] paramsArray = new object[5];
			paramsArray[0] = newWb;
			paramsArray[1] = newMap;
			paramsArray[2] = newUrl;
			paramsArray[3] = newIsRefresh;
			paramsArray.SetValue(cancel, 4);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[4];
		}

		public void WorkbookAfterXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object isRefresh, [In] object result)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkbookAfterXmlImport");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb, map, isRefresh, result);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			NetOffice.ExcelApi.XmlMap newMap = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, map) as NetOffice.ExcelApi.XmlMap;
			bool newIsRefresh = (bool)isRefresh;
			NetOffice.ExcelApi.Enums.XlXmlImportResult newResult = (NetOffice.ExcelApi.Enums.XlXmlImportResult)result;
			object[] paramsArray = new object[4];
			paramsArray[0] = newWb;
			paramsArray[1] = newMap;
			paramsArray[2] = newIsRefresh;
			paramsArray[3] = newResult;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WorkbookBeforeXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkbookBeforeXmlExport");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb, map, url, cancel);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			NetOffice.ExcelApi.XmlMap newMap = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, map) as NetOffice.ExcelApi.XmlMap;
			string newUrl = (string)url;
			object[] paramsArray = new object[4];
			paramsArray[0] = newWb;
			paramsArray[1] = newMap;
			paramsArray[2] = newUrl;
			paramsArray.SetValue(cancel, 3);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[3];
		}

		public void WorkbookAfterXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object result)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkbookAfterXmlExport");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb, map, url, result);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			NetOffice.ExcelApi.XmlMap newMap = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, map) as NetOffice.ExcelApi.XmlMap;
			string newUrl = (string)url;
			NetOffice.ExcelApi.Enums.XlXmlExportResult newResult = (NetOffice.ExcelApi.Enums.XlXmlExportResult)result;
			object[] paramsArray = new object[4];
			paramsArray[0] = newWb;
			paramsArray[1] = newMap;
			paramsArray[2] = newUrl;
			paramsArray[3] = newResult;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WorkbookRowsetComplete([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object description, [In] object sheet, [In] object success)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkbookRowsetComplete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb, description, sheet, success);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			string newDescription = (string)description;
			string newSheet = (string)sheet;
			bool newSuccess = (bool)success;
			object[] paramsArray = new object[4];
			paramsArray[0] = newWb;
			paramsArray[1] = newDescription;
			paramsArray[2] = newSheet;
			paramsArray[3] = newSuccess;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void AfterCalculate()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterCalculate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
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
			NetOffice.ExcelApi.PivotTable newTargetPivotTable = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as NetOffice.ExcelApi.PivotTable;
			NetOffice.ExcelApi.Range newTargetRange = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, targetRange) as NetOffice.ExcelApi.Range;
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
			NetOffice.ExcelApi.PivotTable newTargetPivotTable = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as NetOffice.ExcelApi.PivotTable;
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
			NetOffice.ExcelApi.PivotTable newTargetPivotTable = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as NetOffice.ExcelApi.PivotTable;
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
			NetOffice.ExcelApi.PivotTable newTargetPivotTable = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as NetOffice.ExcelApi.PivotTable;
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

		public void ProtectedViewWindowOpen([In, MarshalAs(UnmanagedType.IDispatch)] object pvw)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProtectedViewWindowOpen");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pvw);
				return;
			}

			NetOffice.ExcelApi.ProtectedViewWindow newPvw = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pvw) as NetOffice.ExcelApi.ProtectedViewWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPvw;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void ProtectedViewWindowBeforeEdit([In, MarshalAs(UnmanagedType.IDispatch)] object pvw, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProtectedViewWindowBeforeEdit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pvw, cancel);
				return;
			}

			NetOffice.ExcelApi.ProtectedViewWindow newPvw = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pvw) as NetOffice.ExcelApi.ProtectedViewWindow;
			object[] paramsArray = new object[2];
			paramsArray[0] = newPvw;
			paramsArray.SetValue(cancel, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void ProtectedViewWindowBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object pvw, [In] object reason, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProtectedViewWindowBeforeClose");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pvw, reason, cancel);
				return;
			}

			NetOffice.ExcelApi.ProtectedViewWindow newPvw = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pvw) as NetOffice.ExcelApi.ProtectedViewWindow;
			NetOffice.ExcelApi.Enums.XlProtectedViewCloseReason newReason = (NetOffice.ExcelApi.Enums.XlProtectedViewCloseReason)reason;
			object[] paramsArray = new object[3];
			paramsArray[0] = newPvw;
			paramsArray[1] = newReason;
			paramsArray.SetValue(cancel, 2);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[2];
		}

		public void ProtectedViewWindowResize([In, MarshalAs(UnmanagedType.IDispatch)] object pvw)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProtectedViewWindowResize");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pvw);
				return;
			}

			NetOffice.ExcelApi.ProtectedViewWindow newPvw = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pvw) as NetOffice.ExcelApi.ProtectedViewWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPvw;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void ProtectedViewWindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvw)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProtectedViewWindowActivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pvw);
				return;
			}

			NetOffice.ExcelApi.ProtectedViewWindow newPvw = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pvw) as NetOffice.ExcelApi.ProtectedViewWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPvw;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void ProtectedViewWindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object pvw)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ProtectedViewWindowDeactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(pvw);
				return;
			}

			NetOffice.ExcelApi.ProtectedViewWindow newPvw = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, pvw) as NetOffice.ExcelApi.ProtectedViewWindow;
			object[] paramsArray = new object[1];
			paramsArray[0] = newPvw;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WorkbookAfterSave([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In] object success)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkbookAfterSave");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb, success);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			bool newSuccess = (bool)success;
			object[] paramsArray = new object[2];
			paramsArray[0] = newWb;
			paramsArray[1] = newSuccess;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void WorkbookNewChart([In, MarshalAs(UnmanagedType.IDispatch)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object ch)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WorkbookNewChart");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wb, ch);
				return;
			}

			NetOffice.ExcelApi.Workbook newWb = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, wb) as NetOffice.ExcelApi.Workbook;
			NetOffice.ExcelApi.Chart newCh = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, ch) as NetOffice.ExcelApi.Chart;
			object[] paramsArray = new object[2];
			paramsArray[0] = newWb;
			paramsArray[1] = newCh;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}