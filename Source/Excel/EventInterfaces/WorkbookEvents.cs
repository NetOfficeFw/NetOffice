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
	[ComImport, Guid("00024412-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface WorkbookEvents
	{
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(682)]
		void Open();

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(304)]
		void Activate();

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1530)]
		void Deactivate();

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1546)]
		void BeforeClose([In] [Out] ref object cancel);

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1547)]
		void BeforeSave([In] object saveAsUI, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1549)]
		void BeforePrint([In] [Out] ref object cancel);

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1550)]
		void NewSheet([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1552)]
		void AddinInstall();

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1553)]
		void AddinUninstall();

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1554)]
		void WindowResize([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1556)]
		void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1557)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn);

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1558)]
		void SheetSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1559)]
		void SheetBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1560)]
		void SheetBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1561)]
		void SheetActivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1562)]
		void SheetDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1563)]
		void SheetCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1564)]
		void SheetChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1854)]
		void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2157)]
		void SheetPivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2158)]
		void PivotTableCloseConnection([In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2159)]
		void PivotTableOpenConnection([In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2266)]
		void Sync([In] object syncEventType);

		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2283)]
		void BeforeXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object isRefresh, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2285)]
		void AfterXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object isRefresh, [In] object result);

		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2287)]
		void BeforeXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2288)]
		void AfterXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object result);

		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2610)]
		void RowsetComplete([In] object description, [In] object sheet, [In] object success);

		[SupportByVersionAttribute("Excel", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2895)]
		void SheetPivotTableAfterValueChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In, MarshalAs(UnmanagedType.IDispatch)] object targetRange);

		[SupportByVersionAttribute("Excel", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2896)]
		void SheetPivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Excel", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2897)]
		void SheetPivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel);

		[SupportByVersionAttribute("Excel", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2898)]
		void SheetPivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd);

		[SupportByVersionAttribute("Excel", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2899)]
		void SheetPivotTableChangeSync([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersionAttribute("Excel", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2900)]
		void AfterSave([In] object success);

		[SupportByVersionAttribute("Excel", 14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2901)]
		void NewChart([In, MarshalAs(UnmanagedType.IDispatch)] object ch);

		[SupportByVersionAttribute("Excel", 15, 16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3073)]
		void SheetLensGalleryRenderComplete([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[SupportByVersionAttribute("Excel", 15, 16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3074)]
		void SheetTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target);

		[SupportByVersionAttribute("Excel", 15, 16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3075)]
		void ModelChange([In, MarshalAs(UnmanagedType.IDispatch)] object changes);

		[SupportByVersionAttribute("Excel", 15, 16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3077)]
		void SheetBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object sh);
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
			_eventBinding.RaiseCustomEvent("Open", ref paramsArray);
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
			_eventBinding.RaiseCustomEvent("BeforeClose", ref paramsArray);

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

			bool newSaveAsUI = Convert.ToBoolean(saveAsUI);
			object[] paramsArray = new object[2];
			paramsArray[0] = newSaveAsUI;
			paramsArray.SetValue(cancel, 1);
			_eventBinding.RaiseCustomEvent("BeforeSave", ref paramsArray);

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
			_eventBinding.RaiseCustomEvent("BeforePrint", ref paramsArray);

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

			object newSh = Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSh;
			_eventBinding.RaiseCustomEvent("NewSheet", ref paramsArray);
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
			_eventBinding.RaiseCustomEvent("AddinInstall", ref paramsArray);
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
			_eventBinding.RaiseCustomEvent("AddinUninstall", ref paramsArray);
		}

		public void WindowResize([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowResize");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wn);
				return;
			}

			NetOffice.ExcelApi.Window newWn = Factory.CreateObjectFromComProxy(_eventClass, wn) as NetOffice.ExcelApi.Window;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			_eventBinding.RaiseCustomEvent("WindowResize", ref paramsArray);
		}

		public void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object wn)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("WindowActivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(wn);
				return;
			}

			NetOffice.ExcelApi.Window newWn = Factory.CreateObjectFromComProxy(_eventClass, wn) as NetOffice.ExcelApi.Window;
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

			NetOffice.ExcelApi.Window newWn = Factory.CreateObjectFromComProxy(_eventClass, wn) as NetOffice.ExcelApi.Window;
			object[] paramsArray = new object[1];
			paramsArray[0] = newWn;
			_eventBinding.RaiseCustomEvent("WindowDeactivate", ref paramsArray);
		}

		public void SheetSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetSelectionChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, target);
				return;
			}

			object newSh = Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			NetOffice.ExcelApi.Range newTarget = Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Range;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSh;
			paramsArray[1] = newTarget;
			_eventBinding.RaiseCustomEvent("SheetSelectionChange", ref paramsArray);
		}

		public void SheetBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetBeforeDoubleClick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, target, cancel);
				return;
			}

			object newSh = Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			NetOffice.ExcelApi.Range newTarget = Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Range;
			object[] paramsArray = new object[3];
			paramsArray[0] = newSh;
			paramsArray[1] = newTarget;
			paramsArray.SetValue(cancel, 2);
			_eventBinding.RaiseCustomEvent("SheetBeforeDoubleClick", ref paramsArray);

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

			object newSh = Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			NetOffice.ExcelApi.Range newTarget = Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Range;
			object[] paramsArray = new object[3];
			paramsArray[0] = newSh;
			paramsArray[1] = newTarget;
			paramsArray.SetValue(cancel, 2);
			_eventBinding.RaiseCustomEvent("SheetBeforeRightClick", ref paramsArray);

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

			object newSh = Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSh;
			_eventBinding.RaiseCustomEvent("SheetActivate", ref paramsArray);
		}

		public void SheetDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetDeactivate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh);
				return;
			}

			object newSh = Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSh;
			_eventBinding.RaiseCustomEvent("SheetDeactivate", ref paramsArray);
		}

		public void SheetCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetCalculate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh);
				return;
			}

			object newSh = Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSh;
			_eventBinding.RaiseCustomEvent("SheetCalculate", ref paramsArray);
		}

		public void SheetChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, target);
				return;
			}

			object newSh = Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			NetOffice.ExcelApi.Range newTarget = Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Range;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSh;
			paramsArray[1] = newTarget;
			_eventBinding.RaiseCustomEvent("SheetChange", ref paramsArray);
		}

		public void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetFollowHyperlink");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, target);
				return;
			}

			object newSh = Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			NetOffice.ExcelApi.Hyperlink newTarget = Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.Hyperlink;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSh;
			paramsArray[1] = newTarget;
			_eventBinding.RaiseCustomEvent("SheetFollowHyperlink", ref paramsArray);
		}

		public void SheetPivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetPivotTableUpdate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, target);
				return;
			}

			object newSh = Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			NetOffice.ExcelApi.PivotTable newTarget = Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.PivotTable;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSh;
			paramsArray[1] = newTarget;
			_eventBinding.RaiseCustomEvent("SheetPivotTableUpdate", ref paramsArray);
		}

		public void PivotTableCloseConnection([In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PivotTableCloseConnection");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(target);
				return;
			}

			NetOffice.ExcelApi.PivotTable newTarget = Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.PivotTable;
			object[] paramsArray = new object[1];
			paramsArray[0] = newTarget;
			_eventBinding.RaiseCustomEvent("PivotTableCloseConnection", ref paramsArray);
		}

		public void PivotTableOpenConnection([In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PivotTableOpenConnection");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(target);
				return;
			}

			NetOffice.ExcelApi.PivotTable newTarget = Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.PivotTable;
			object[] paramsArray = new object[1];
			paramsArray[0] = newTarget;
			_eventBinding.RaiseCustomEvent("PivotTableOpenConnection", ref paramsArray);
		}

		public void Sync([In] object syncEventType)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Sync");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(syncEventType);
				return;
			}

			NetOffice.OfficeApi.Enums.MsoSyncEventType newSyncEventType = (NetOffice.OfficeApi.Enums.MsoSyncEventType)syncEventType;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSyncEventType;
			_eventBinding.RaiseCustomEvent("Sync", ref paramsArray);
		}

		public void BeforeXmlImport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] object isRefresh, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeXmlImport");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(map, url, isRefresh, cancel);
				return;
			}

			NetOffice.ExcelApi.XmlMap newMap = Factory.CreateObjectFromComProxy(_eventClass, map) as NetOffice.ExcelApi.XmlMap;
			string newUrl = Convert.ToString(url);
			bool newIsRefresh = Convert.ToBoolean(isRefresh);
			object[] paramsArray = new object[4];
			paramsArray[0] = newMap;
			paramsArray[1] = newUrl;
			paramsArray[2] = newIsRefresh;
			paramsArray.SetValue(cancel, 3);
			_eventBinding.RaiseCustomEvent("BeforeXmlImport", ref paramsArray);

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

			NetOffice.ExcelApi.XmlMap newMap = Factory.CreateObjectFromComProxy(_eventClass, map) as NetOffice.ExcelApi.XmlMap;
			bool newIsRefresh = Convert.ToBoolean(isRefresh);
			NetOffice.ExcelApi.Enums.XlXmlImportResult newResult = (NetOffice.ExcelApi.Enums.XlXmlImportResult)result;
			object[] paramsArray = new object[3];
			paramsArray[0] = newMap;
			paramsArray[1] = newIsRefresh;
			paramsArray[2] = newResult;
			_eventBinding.RaiseCustomEvent("AfterXmlImport", ref paramsArray);
		}

		public void BeforeXmlExport([In, MarshalAs(UnmanagedType.IDispatch)] object map, [In] object url, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeXmlExport");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(map, url, cancel);
				return;
			}

			NetOffice.ExcelApi.XmlMap newMap = Factory.CreateObjectFromComProxy(_eventClass, map) as NetOffice.ExcelApi.XmlMap;
			string newUrl = Convert.ToString(url);
			object[] paramsArray = new object[3];
			paramsArray[0] = newMap;
			paramsArray[1] = newUrl;
			paramsArray.SetValue(cancel, 2);
			_eventBinding.RaiseCustomEvent("BeforeXmlExport", ref paramsArray);

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

			NetOffice.ExcelApi.XmlMap newMap = Factory.CreateObjectFromComProxy(_eventClass, map) as NetOffice.ExcelApi.XmlMap;
			string newUrl = Convert.ToString(url);
			NetOffice.ExcelApi.Enums.XlXmlExportResult newResult = (NetOffice.ExcelApi.Enums.XlXmlExportResult)result;
			object[] paramsArray = new object[3];
			paramsArray[0] = newMap;
			paramsArray[1] = newUrl;
			paramsArray[2] = newResult;
			_eventBinding.RaiseCustomEvent("AfterXmlExport", ref paramsArray);
		}

		public void RowsetComplete([In] object description, [In] object sheet, [In] object success)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("RowsetComplete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(description, sheet, success);
				return;
			}

			string newDescription = Convert.ToString(description);
			string newSheet = Convert.ToString(sheet);
			bool newSuccess = Convert.ToBoolean(success);
			object[] paramsArray = new object[3];
			paramsArray[0] = newDescription;
			paramsArray[1] = newSheet;
			paramsArray[2] = newSuccess;
			_eventBinding.RaiseCustomEvent("RowsetComplete", ref paramsArray);
		}

		public void SheetPivotTableAfterValueChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In, MarshalAs(UnmanagedType.IDispatch)] object targetRange)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetPivotTableAfterValueChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, targetPivotTable, targetRange);
				return;
			}

			object newSh = Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as NetOffice.ExcelApi.PivotTable;
			NetOffice.ExcelApi.Range newTargetRange = Factory.CreateObjectFromComProxy(_eventClass, targetRange) as NetOffice.ExcelApi.Range;
			object[] paramsArray = new object[3];
			paramsArray[0] = newSh;
			paramsArray[1] = newTargetPivotTable;
			paramsArray[2] = newTargetRange;
			_eventBinding.RaiseCustomEvent("SheetPivotTableAfterValueChange", ref paramsArray);
		}

		public void SheetPivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object targetPivotTable, [In] object valueChangeStart, [In] object valueChangeEnd, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetPivotTableBeforeAllocateChanges");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, targetPivotTable, valueChangeStart, valueChangeEnd, cancel);
				return;
			}

			object newSh = Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as NetOffice.ExcelApi.PivotTable;
			Int32 newValueChangeStart = Convert.ToInt32(valueChangeStart);
			Int32 newValueChangeEnd = Convert.ToInt32(valueChangeEnd);
			object[] paramsArray = new object[5];
			paramsArray[0] = newSh;
			paramsArray[1] = newTargetPivotTable;
			paramsArray[2] = newValueChangeStart;
			paramsArray[3] = newValueChangeEnd;
			paramsArray.SetValue(cancel, 4);
			_eventBinding.RaiseCustomEvent("SheetPivotTableBeforeAllocateChanges", ref paramsArray);

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

			object newSh = Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as NetOffice.ExcelApi.PivotTable;
			Int32 newValueChangeStart = Convert.ToInt32(valueChangeStart);
			Int32 newValueChangeEnd = Convert.ToInt32(valueChangeEnd);
			object[] paramsArray = new object[5];
			paramsArray[0] = newSh;
			paramsArray[1] = newTargetPivotTable;
			paramsArray[2] = newValueChangeStart;
			paramsArray[3] = newValueChangeEnd;
			paramsArray.SetValue(cancel, 4);
			_eventBinding.RaiseCustomEvent("SheetPivotTableBeforeCommitChanges", ref paramsArray);

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

			object newSh = Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			NetOffice.ExcelApi.PivotTable newTargetPivotTable = Factory.CreateObjectFromComProxy(_eventClass, targetPivotTable) as NetOffice.ExcelApi.PivotTable;
			Int32 newValueChangeStart = Convert.ToInt32(valueChangeStart);
			Int32 newValueChangeEnd = Convert.ToInt32(valueChangeEnd);
			object[] paramsArray = new object[4];
			paramsArray[0] = newSh;
			paramsArray[1] = newTargetPivotTable;
			paramsArray[2] = newValueChangeStart;
			paramsArray[3] = newValueChangeEnd;
			_eventBinding.RaiseCustomEvent("SheetPivotTableBeforeDiscardChanges", ref paramsArray);
		}

		public void SheetPivotTableChangeSync([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetPivotTableChangeSync");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, target);
				return;
			}

			object newSh = Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			NetOffice.ExcelApi.PivotTable newTarget = Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.PivotTable;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSh;
			paramsArray[1] = newTarget;
			_eventBinding.RaiseCustomEvent("SheetPivotTableChangeSync", ref paramsArray);
		}

		public void AfterSave([In] object success)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterSave");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(success);
				return;
			}

			bool newSuccess = Convert.ToBoolean(success);
			object[] paramsArray = new object[1];
			paramsArray[0] = newSuccess;
			_eventBinding.RaiseCustomEvent("AfterSave", ref paramsArray);
		}

		public void NewChart([In, MarshalAs(UnmanagedType.IDispatch)] object ch)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("NewChart");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(ch);
				return;
			}

			NetOffice.ExcelApi.Chart newCh = Factory.CreateObjectFromComProxy(_eventClass, ch) as NetOffice.ExcelApi.Chart;
			object[] paramsArray = new object[1];
			paramsArray[0] = newCh;
			_eventBinding.RaiseCustomEvent("NewChart", ref paramsArray);
		}

		public void SheetLensGalleryRenderComplete([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetLensGalleryRenderComplete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh);
				return;
			}

			object newSh = Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSh;
			_eventBinding.RaiseCustomEvent("SheetLensGalleryRenderComplete", ref paramsArray);
		}

		public void SheetTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.IDispatch)] object target)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetTableUpdate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh, target);
				return;
			}

			object newSh = Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			NetOffice.ExcelApi.TableObject newTarget = Factory.CreateObjectFromComProxy(_eventClass, target) as NetOffice.ExcelApi.TableObject;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSh;
			paramsArray[1] = newTarget;
			_eventBinding.RaiseCustomEvent("SheetTableUpdate", ref paramsArray);
		}

		public void ModelChange([In, MarshalAs(UnmanagedType.IDispatch)] object changes)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ModelChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(changes);
				return;
			}

			NetOffice.ExcelApi.ModelChanges newChanges = Factory.CreateObjectFromComProxy(_eventClass, changes) as NetOffice.ExcelApi.ModelChanges;
			object[] paramsArray = new object[1];
			paramsArray[0] = newChanges;
			_eventBinding.RaiseCustomEvent("ModelChange", ref paramsArray);
		}

		public void SheetBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SheetBeforeDelete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(sh);
				return;
			}

			object newSh = Factory.CreateObjectFromComProxy(_eventClass, sh) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newSh;
			_eventBinding.RaiseCustomEvent("SheetBeforeDelete", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}