using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.ExcelApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
	[ComImport, Guid("00024411-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface DocEvents
	{
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1543)]
		void SelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1537)]
		void BeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1534)]
		void BeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(304)]
		void Activate();

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1530)]
		void Deactivate();

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(279)]
		void Calculate();

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1545)]
		void Change([In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1470)]
		void FollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2156)]
		void PivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersionAttribute("Excel", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2886)]
		void PivotTableAfterValueChange([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In, MarshalAs(UnmanagedType.IDispatch)] object targetRange);

		[SupportByVersionAttribute("Excel", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2889)]
		void PivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Excel", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2892)]
		void PivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Excel", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2893)]
		void PivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd);

		[SupportByVersionAttribute("Excel", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2894)]
		void PivotTableChangeSync([In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersionAttribute("Excel", 15, 16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3070)]
		void LensGalleryRenderComplete();

		[SupportByVersionAttribute("Excel", 15, 16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3071)]
		void TableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersionAttribute("Excel", 15, 16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3072)]
		void BeforeDelete();
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

		#region DocEvents Members
		
		public void SelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SelectionChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(target);
				return;
			}

			NetOffice.ExcelApi.Range newTarget = Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Range;
			object[] paramsArray = new object[1];
			paramsArray[0] = newTarget;
			_eventBinding.RaiseCustomEvent("SelectionChange", ref paramsArray);
		}

		public void BeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeDoubleClick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(target, cancel);
				return;
			}

			NetOffice.ExcelApi.Range newTarget = Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Range;
			object[] paramsArray = new object[2];
			paramsArray[0] = newTarget;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("BeforeDoubleClick", ref paramsArray);

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

			NetOffice.ExcelApi.Range newTarget = Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Range;
			object[] paramsArray = new object[2];
			paramsArray[0] = newTarget;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("BeforeRightClick", ref paramsArray);

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
			_eventBinding.RaiseCustomEvent("Activate", ref paramsArray);
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
			_eventBinding.RaiseCustomEvent("Deactivate", ref paramsArray);
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
			_eventBinding.RaiseCustomEvent("Calculate", ref paramsArray);
		}

		public void Change([In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Change");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(target);
				return;
			}

			NetOffice.ExcelApi.Range newTarget = Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Range;
			object[] paramsArray = new object[1];
			paramsArray[0] = newTarget;
			_eventBinding.RaiseCustomEvent("Change", ref paramsArray);
		}

		public void FollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("FollowHyperlink");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(target);
				return;
			}

			NetOffice.ExcelApi.Hyperlink newTarget = Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Hyperlink;
			object[] paramsArray = new object[1];
			paramsArray[0] = newTarget;
			_eventBinding.RaiseCustomEvent("FollowHyperlink", ref paramsArray);
		}

		public void PivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PivotTableUpdate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(target);
				return;
			}

			NetOffice.ExcelApi.PivotTable newTarget = Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.PivotTable;
			object[] paramsArray = new object[1];
			paramsArray[0] = newTarget;
			_eventBinding.RaiseCustomEvent("PivotTableUpdate", ref paramsArray);
		}

		public void PivotTableAfterValueChange([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In, MarshalAs(UnmanagedType.IDispatch)] object targetRange)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PivotTableAfterValueChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(targetPivotTable, targetRange);
				return;
			}

			NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as NetOffice.ExcelApi.PivotTable;
			NetOffice.ExcelApi.Range newTargetRange = Factory.CreateObjectFromComProxy(_eventClass, targetRange) as NetOffice.ExcelApi.Range;
			object[] paramsArray = new object[2];
			paramsArray[0] = newTargetPivotTable;
			paramsArray[1] = newTargetRange;
			_eventBinding.RaiseCustomEvent("PivotTableAfterValueChange", ref paramsArray);
		}

		public void PivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PivotTableBeforeAllocateChanges");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(targetPivotTable, valueChangeStart, valueChangeEnd, cancel);
				return;
			}

			NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as NetOffice.ExcelApi.PivotTable;
			Int32 newValueChangeStart = Convert.ToInt32(valueChangeStart);
			Int32 newValueChangeEnd = Convert.ToInt32(valueChangeEnd);
			object[] paramsArray = new object[4];
			paramsArray[0] = newTargetPivotTable;
			paramsArray[1] = newValueChangeStart;
			paramsArray[2] = newValueChangeEnd;
			paramsArray.SetValue(cancel, 3);
			_eventBinding.RaiseCustomEvent("PivotTableBeforeAllocateChanges", ref paramsArray);

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

			NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as NetOffice.ExcelApi.PivotTable;
			Int32 newValueChangeStart = Convert.ToInt32(valueChangeStart);
			Int32 newValueChangeEnd = Convert.ToInt32(valueChangeEnd);
			object[] paramsArray = new object[4];
			paramsArray[0] = newTargetPivotTable;
			paramsArray[1] = newValueChangeStart;
			paramsArray[2] = newValueChangeEnd;
			paramsArray.SetValue(cancel, 3);
			_eventBinding.RaiseCustomEvent("PivotTableBeforeCommitChanges", ref paramsArray);

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

			NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as NetOffice.ExcelApi.PivotTable;
			Int32 newValueChangeStart = Convert.ToInt32(valueChangeStart);
			Int32 newValueChangeEnd = Convert.ToInt32(valueChangeEnd);
			object[] paramsArray = new object[3];
			paramsArray[0] = newTargetPivotTable;
			paramsArray[1] = newValueChangeStart;
			paramsArray[2] = newValueChangeEnd;
			_eventBinding.RaiseCustomEvent("PivotTableBeforeDiscardChanges", ref paramsArray);
		}

		public void PivotTableChangeSync([In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PivotTableChangeSync");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(target);
				return;
			}

			NetOffice.ExcelApi.PivotTable newTarget = Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.PivotTable;
			object[] paramsArray = new object[1];
			paramsArray[0] = newTarget;
			_eventBinding.RaiseCustomEvent("PivotTableChangeSync", ref paramsArray);
		}

		public void LensGalleryRenderComplete()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("LensGalleryRenderComplete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("LensGalleryRenderComplete", ref paramsArray);
		}

		public void TableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("TableUpdate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(target);
				return;
			}

			NetOffice.ExcelApi.TableObject newTarget = Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.TableObject;
			object[] paramsArray = new object[1];
			paramsArray[0] = newTarget;
			_eventBinding.RaiseCustomEvent("TableUpdate", ref paramsArray);
		}

		public void BeforeDelete()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("BeforeDelete", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}