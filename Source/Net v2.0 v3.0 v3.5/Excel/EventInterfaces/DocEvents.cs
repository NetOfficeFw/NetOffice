using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace NetOffice.ExcelApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
	[ComImport, Guid("00024411-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface DocEvents
	{
		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1543)]
		void SelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1537)]
		void BeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1534)]
		void BeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(304)]
		void Activate();

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1530)]
		void Deactivate();

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(279)]
		void Calculate();

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1545)]
		void Change([In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1470)]
		void FollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2156)]
		void PivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByLibrary("XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2886)]
		void PivotTableAfterValueChange([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In, MarshalAs(UnmanagedType.IDispatch)] object targetRange);

		[SupportByLibrary("XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2889)]
		void PivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel);

		[SupportByLibrary("XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2892)]
		void PivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel);

		[SupportByLibrary("XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2893)]
		void PivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd);

		[SupportByLibrary("XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2894)]
		void PivotTableChangeSync([In, MarshalAs(UnmanagedType.IDispatch)] object target);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class DocEvents_SinkHelper : SinkHelper, DocEvents
	{
		#region Static
		
		public static readonly string Id = "00024411-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public DocEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region DocEvents Members
		
		public void SelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SelectionChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(target);
				return;
			}

			NetOffice.ExcelApi.Range newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Range;
			object[] paramsArray = new object[1];
			paramsArray[0] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void BeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeDoubleClick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(target, cancel);
				return;
			}

			NetOffice.ExcelApi.Range newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Range;
			object[] paramsArray = new object[2];
			paramsArray[0] = newTarget;
			paramsArray.SetValue(cancel, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[1];
		}

		public void BeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeRightClick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(target, cancel);
				return;
			}

			NetOffice.ExcelApi.Range newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Range;
			object[] paramsArray = new object[2];
			paramsArray[0] = newTarget;
			paramsArray.SetValue(cancel, 1);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[1];
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

		public void Calculate()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Calculate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void Change([In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Change");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(target);
				return;
			}

			NetOffice.ExcelApi.Range newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Range;
			object[] paramsArray = new object[1];
			paramsArray[0] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void FollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("FollowHyperlink");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(target);
				return;
			}

			NetOffice.ExcelApi.Hyperlink newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Hyperlink;
			object[] paramsArray = new object[1];
			paramsArray[0] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void PivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PivotTableUpdate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(target);
				return;
			}

			NetOffice.ExcelApi.PivotTable newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.PivotTable;
			object[] paramsArray = new object[1];
			paramsArray[0] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void PivotTableAfterValueChange([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In, MarshalAs(UnmanagedType.IDispatch)] object targetRange)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PivotTableAfterValueChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(targetPivotTable, targetRange);
				return;
			}

			NetOffice.ExcelApi.PivotTable newTargetPivotTable = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as NetOffice.ExcelApi.PivotTable;
			NetOffice.ExcelApi.Range newTargetRange = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, targetRange) as NetOffice.ExcelApi.Range;
			object[] paramsArray = new object[2];
			paramsArray[0] = newTargetPivotTable;
			paramsArray[1] = newTargetRange;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void PivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PivotTableBeforeAllocateChanges");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(targetPivotTable, valueChangeStart, valueChangeEnd, cancel);
				return;
			}

			NetOffice.ExcelApi.PivotTable newTargetPivotTable = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as NetOffice.ExcelApi.PivotTable;
			Int32 newValueChangeStart = (Int32)valueChangeStart;
			Int32 newValueChangeEnd = (Int32)valueChangeEnd;
			object[] paramsArray = new object[4];
			paramsArray[0] = newTargetPivotTable;
			paramsArray[1] = newValueChangeStart;
			paramsArray[2] = newValueChangeEnd;
			paramsArray.SetValue(cancel, 3);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[3];
		}

		public void PivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PivotTableBeforeCommitChanges");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(targetPivotTable, valueChangeStart, valueChangeEnd, cancel);
				return;
			}

			NetOffice.ExcelApi.PivotTable newTargetPivotTable = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as NetOffice.ExcelApi.PivotTable;
			Int32 newValueChangeStart = (Int32)valueChangeStart;
			Int32 newValueChangeEnd = (Int32)valueChangeEnd;
			object[] paramsArray = new object[4];
			paramsArray[0] = newTargetPivotTable;
			paramsArray[1] = newValueChangeStart;
			paramsArray[2] = newValueChangeEnd;
			paramsArray.SetValue(cancel, 3);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[3];
		}

		public void PivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PivotTableBeforeDiscardChanges");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(targetPivotTable, valueChangeStart, valueChangeEnd);
				return;
			}

			NetOffice.ExcelApi.PivotTable newTargetPivotTable = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as NetOffice.ExcelApi.PivotTable;
			Int32 newValueChangeStart = (Int32)valueChangeStart;
			Int32 newValueChangeEnd = (Int32)valueChangeEnd;
			object[] paramsArray = new object[3];
			paramsArray[0] = newTargetPivotTable;
			paramsArray[1] = newValueChangeStart;
			paramsArray[2] = newValueChangeEnd;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void PivotTableChangeSync([In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PivotTableChangeSync");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(target);
				return;
			}

			NetOffice.ExcelApi.PivotTable newTarget = LateBindingApi.Core.Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.PivotTable;
			object[] paramsArray = new object[1];
			paramsArray[0] = newTarget;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}