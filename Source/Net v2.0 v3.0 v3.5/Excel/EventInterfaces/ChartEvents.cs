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
	[ComImport, Guid("0002440F-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ChartEvents
	{
		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(304)]
		void Activate();

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1530)]
		void Deactivate();

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(256)]
		void Resize();

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1531)]
		void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1532)]
		void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1533)]
		void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1534)]
		void BeforeRightClick([In] [Out] ref object cancel);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1535)]
		void DragPlot();

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1536)]
		void DragOver();

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1537)]
		void BeforeDoubleClick([In] object elementID, [In] object arg1, [In] object arg2, [In] [Out] ref object cancel);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(235)]
		void Select([In] object elementID, [In] object arg1, [In] object arg2);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1538)]
		void SeriesChange([In] object seriesIndex, [In] object pointIndex);

		[SupportByLibrary("XL09","XL10","XL11","XL12","XL14")]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(279)]
		void Calculate();
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ChartEvents_SinkHelper : SinkHelper, ChartEvents
	{
		#region Static
		
		public static readonly string Id = "0002440F-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public ChartEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(connectPoint);
		}
		
		#endregion
		
		#region ChartEvents Members
		
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

		public void Resize()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Resize");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseDown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, shift, x, y);
				return;
			}

			Int32 newButton = (Int32)button;
			Int32 newShift = (Int32)shift;
			Int32 newx = (Int32)x;
			Int32 newy = (Int32)y;
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseUp");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, shift, x, y);
				return;
			}

			Int32 newButton = (Int32)button;
			Int32 newShift = (Int32)shift;
			Int32 newx = (Int32)x;
			Int32 newy = (Int32)y;
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseMove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, shift, x, y);
				return;
			}

			Int32 newButton = (Int32)button;
			Int32 newShift = (Int32)shift;
			Int32 newx = (Int32)x;
			Int32 newy = (Int32)y;
			object[] paramsArray = new object[4];
			paramsArray[0] = newButton;
			paramsArray[1] = newShift;
			paramsArray[2] = newx;
			paramsArray[3] = newy;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void BeforeRightClick([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeRightClick");
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

		public void DragPlot()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DragPlot");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void DragOver()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DragOver");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void BeforeDoubleClick([In] object elementID, [In] object arg1, [In] object arg2, [In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeDoubleClick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(elementID, arg1, arg2, cancel);
				return;
			}

			Int32 newElementID = (Int32)elementID;
			Int32 newArg1 = (Int32)arg1;
			Int32 newArg2 = (Int32)arg2;
			object[] paramsArray = new object[4];
			paramsArray[0] = newElementID;
			paramsArray[1] = newArg1;
			paramsArray[2] = newArg2;
			paramsArray.SetValue(cancel, 3);
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);

			cancel = (bool)paramsArray[3];
		}

		public void Select([In] object elementID, [In] object arg1, [In] object arg2)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Select");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(elementID, arg1, arg2);
				return;
			}

			Int32 newElementID = (Int32)elementID;
			Int32 newArg1 = (Int32)arg1;
			Int32 newArg2 = (Int32)arg2;
			object[] paramsArray = new object[3];
			paramsArray[0] = newElementID;
			paramsArray[1] = newArg1;
			paramsArray[2] = newArg2;
			foreach(Delegate delItem in recipients)
				delItem.Method.Invoke(delItem.Target, paramsArray);
		}

		public void SeriesChange([In] object seriesIndex, [In] object pointIndex)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SeriesChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(seriesIndex, pointIndex);
				return;
			}

			Int32 newSeriesIndex = (Int32)seriesIndex;
			Int32 newPointIndex = (Int32)pointIndex;
			object[] paramsArray = new object[2];
			paramsArray[0] = newSeriesIndex;
			paramsArray[1] = newPointIndex;
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

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}