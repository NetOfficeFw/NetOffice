using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;

namespace NetOffice.AccessApi
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
	[ComImport, Guid("331FDCFB-CF31-11CD-8701-00AA003F0F07"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _FormEvents
	{
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2067)]
		void Load();

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2058)]
		void Current();

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2059)]
		void BeforeInsert([In] [Out] ref object cancel);

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2060)]
		void AfterInsert();

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2061)]
		void BeforeUpdate([In] [Out] ref object cancel);

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2062)]
		void AfterUpdate();

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2063)]
		void Delete([In] [Out] ref object cancel);

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2064)]
		void BeforeDelConfirm([In] [Out] ref object cancel, [In] [Out] ref object response);

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2065)]
		void AfterDelConfirm([In] [Out] ref object status);

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2066)]
		void Open([In] [Out] ref object cancel);

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2068)]
		void Resize();

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2069)]
		void Unload([In] [Out] ref object cancel);

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2070)]
		void Close();

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2071)]
		void Activate();

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2072)]
		void Deactivate();

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2073)]
		void GotFocus();

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2074)]
		void LostFocus();

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void Click();

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-601)]
		void DblClick([In] [Out] ref object cancel);

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-605)]
		void MouseDown([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-606)]
		void MouseMove([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-607)]
		void MouseUp([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y);

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-602)]
		void KeyDown([In] [Out] ref object keyCode, [In] [Out] ref object shift);

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-603)]
		void KeyPress([In] [Out] ref object keyAscii);

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-604)]
		void KeyUp([In] [Out] ref object keyCode, [In] [Out] ref object shift);

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2083)]
		void Error([In] [Out] ref object dataErr, [In] [Out] ref object response);

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2084)]
		void Timer();

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2155)]
		void Filter([In] [Out] ref object cancel, [In] [Out] ref object filterType);

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2156)]
		void ApplyFilter([In] [Out] ref object cancel, [In] [Out] ref object applyType);

		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2205)]
		void Dirty([In] [Out] ref object cancel);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2145)]
		void Undo([In] [Out] ref object cancel);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2334)]
		void RecordExit([In] [Out] ref object cancel);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2369)]
		void BeginBatchEdit([In] [Out] ref object cancel);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2370)]
		void UndoBatchEdit([In] [Out] ref object cancel);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2371)]
		void BeforeBeginTransaction([In] [Out] ref object cancel, [In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2372)]
		void AfterBeginTransaction([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2373)]
		void BeforeCommitTransaction([In] [Out] ref object cancel, [In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2374)]
		void AfterCommitTransaction([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2375)]
		void RollbackTransaction([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2383)]
		void OnConnect();

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2384)]
		void OnDisconnect();

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2385)]
		void PivotTableChange([In] object reason);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2386)]
		void Query();

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2387)]
		void BeforeQuery();

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2388)]
		void SelectionChange();

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2389)]
		void CommandBeforeExecute([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2390)]
		void CommandChecked([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object _checked);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2391)]
		void CommandEnabled([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object enabled);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2392)]
		void CommandExecute([In] object command);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2394)]
		void DataSetChange();

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2395)]
		void BeforeScreenTip([In, MarshalAs(UnmanagedType.IDispatch)] object screenTipText, [In, MarshalAs(UnmanagedType.IDispatch)] object sourceObject);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2399)]
		void BeforeRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2397)]
		void AfterRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2396)]
		void AfterFinalRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2398)]
		void AfterLayout([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2401)]
		void MouseWheel([In] object page, [In] object count);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2402)]
		void ViewChange([In] object reason);

		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2403)]
		void DataChange([In] object reason);
	}
	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _FormEvents_SinkHelper : SinkHelper, _FormEvents
	{
		#region Static
		
		public static readonly string Id = "331FDCFB-CF31-11CD-8701-00AA003F0F07";
		
		#endregion
	
		#region Fields

		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
		
		#region Construction

		public _FormEvents_SinkHelper(COMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

		#region _FormEvents Members
		
		public void Load()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Load");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("Load", ref paramsArray);
		}

		public void Current()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Current");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("Current", ref paramsArray);
		}

		public void BeforeInsert([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeInsert");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			_eventBinding.RaiseCustomEvent("BeforeInsert", ref paramsArray);

			cancel = (Int16)paramsArray[0];
		}

		public void AfterInsert()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterInsert");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("AfterInsert", ref paramsArray);
		}

		public void BeforeUpdate([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeUpdate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			_eventBinding.RaiseCustomEvent("BeforeUpdate", ref paramsArray);

			cancel = (Int16)paramsArray[0];
		}

		public void AfterUpdate()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterUpdate");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("AfterUpdate", ref paramsArray);
		}

		public void Delete([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Delete");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			_eventBinding.RaiseCustomEvent("Delete", ref paramsArray);

			cancel = (Int16)paramsArray[0];
		}

		public void BeforeDelConfirm([In] [Out] ref object cancel, [In] [Out] ref object response)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeDelConfirm");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel, response);
				return;
			}

			object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(response, 1);
			_eventBinding.RaiseCustomEvent("BeforeDelConfirm", ref paramsArray);

			cancel = (Int16)paramsArray[0];
			response = (Int16)paramsArray[1];
		}

		public void AfterDelConfirm([In] [Out] ref object status)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterDelConfirm");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(status);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(status, 0);
			_eventBinding.RaiseCustomEvent("AfterDelConfirm", ref paramsArray);

			status = (Int16)paramsArray[0];
		}

		public void Open([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Open");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			_eventBinding.RaiseCustomEvent("Open", ref paramsArray);

			cancel = (Int16)paramsArray[0];
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
			_eventBinding.RaiseCustomEvent("Resize", ref paramsArray);
		}

		public void Unload([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Unload");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			_eventBinding.RaiseCustomEvent("Unload", ref paramsArray);

			cancel = (Int16)paramsArray[0];
		}

		public void Close()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Close");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("Close", ref paramsArray);
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

		public void GotFocus()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("GotFocus");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("GotFocus", ref paramsArray);
		}

		public void LostFocus()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("LostFocus");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("LostFocus", ref paramsArray);
		}

		public void Click()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Click");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("Click", ref paramsArray);
		}

		public void DblClick([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DblClick");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			_eventBinding.RaiseCustomEvent("DblClick", ref paramsArray);

			cancel = (Int16)paramsArray[0];
		}

		public void MouseDown([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseDown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, shift, x, y);
				return;
			}

			object[] paramsArray = new object[4];
			paramsArray.SetValue(button, 0);
			paramsArray.SetValue(shift, 1);
			paramsArray.SetValue(x, 2);
			paramsArray.SetValue(y, 3);
			_eventBinding.RaiseCustomEvent("MouseDown", ref paramsArray);

			button = (Int16)paramsArray[0];
			shift = (Int16)paramsArray[1];
			x = (Single)paramsArray[2];
			y = (Single)paramsArray[3];
		}

		public void MouseMove([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseMove");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, shift, x, y);
				return;
			}

			object[] paramsArray = new object[4];
			paramsArray.SetValue(button, 0);
			paramsArray.SetValue(shift, 1);
			paramsArray.SetValue(x, 2);
			paramsArray.SetValue(y, 3);
			_eventBinding.RaiseCustomEvent("MouseMove", ref paramsArray);

			button = (Int16)paramsArray[0];
			shift = (Int16)paramsArray[1];
			x = (Single)paramsArray[2];
			y = (Single)paramsArray[3];
		}

		public void MouseUp([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseUp");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(button, shift, x, y);
				return;
			}

			object[] paramsArray = new object[4];
			paramsArray.SetValue(button, 0);
			paramsArray.SetValue(shift, 1);
			paramsArray.SetValue(x, 2);
			paramsArray.SetValue(y, 3);
			_eventBinding.RaiseCustomEvent("MouseUp", ref paramsArray);

			button = (Int16)paramsArray[0];
			shift = (Int16)paramsArray[1];
			x = (Single)paramsArray[2];
			y = (Single)paramsArray[3];
		}

		public void KeyDown([In] [Out] ref object keyCode, [In] [Out] ref object shift)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyDown");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyCode, shift);
				return;
			}

			object[] paramsArray = new object[2];
			paramsArray.SetValue(keyCode, 0);
			paramsArray.SetValue(shift, 1);
			_eventBinding.RaiseCustomEvent("KeyDown", ref paramsArray);

			keyCode = (Int16)paramsArray[0];
			shift = (Int16)paramsArray[1];
		}

		public void KeyPress([In] [Out] ref object keyAscii)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyPress");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyAscii);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(keyAscii, 0);
			_eventBinding.RaiseCustomEvent("KeyPress", ref paramsArray);

			keyAscii = (Int16)paramsArray[0];
		}

		public void KeyUp([In] [Out] ref object keyCode, [In] [Out] ref object shift)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("KeyUp");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(keyCode, shift);
				return;
			}

			object[] paramsArray = new object[2];
			paramsArray.SetValue(keyCode, 0);
			paramsArray.SetValue(shift, 1);
			_eventBinding.RaiseCustomEvent("KeyUp", ref paramsArray);

			keyCode = (Int16)paramsArray[0];
			shift = (Int16)paramsArray[1];
		}

		public void Error([In] [Out] ref object dataErr, [In] [Out] ref object response)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Error");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(dataErr, response);
				return;
			}

			object[] paramsArray = new object[2];
			paramsArray.SetValue(dataErr, 0);
			paramsArray.SetValue(response, 1);
			_eventBinding.RaiseCustomEvent("Error", ref paramsArray);

			dataErr = (Int16)paramsArray[0];
			response = (Int16)paramsArray[1];
		}

		public void Timer()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Timer");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("Timer", ref paramsArray);
		}

		public void Filter([In] [Out] ref object cancel, [In] [Out] ref object filterType)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Filter");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel, filterType);
				return;
			}

			object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(filterType, 1);
			_eventBinding.RaiseCustomEvent("Filter", ref paramsArray);

			cancel = (Int16)paramsArray[0];
			filterType = (Int16)paramsArray[1];
		}

		public void ApplyFilter([In] [Out] ref object cancel, [In] [Out] ref object applyType)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ApplyFilter");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel, applyType);
				return;
			}

			object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(applyType, 1);
			_eventBinding.RaiseCustomEvent("ApplyFilter", ref paramsArray);

			cancel = (Int16)paramsArray[0];
			applyType = (Int16)paramsArray[1];
		}

		public void Dirty([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Dirty");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			_eventBinding.RaiseCustomEvent("Dirty", ref paramsArray);

			cancel = (Int16)paramsArray[0];
		}

		public void Undo([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Undo");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			_eventBinding.RaiseCustomEvent("Undo", ref paramsArray);

			cancel = (Int16)paramsArray[0];
		}

		public void RecordExit([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("RecordExit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			_eventBinding.RaiseCustomEvent("RecordExit", ref paramsArray);

			cancel = (Int16)paramsArray[0];
		}

		public void BeginBatchEdit([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeginBatchEdit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			_eventBinding.RaiseCustomEvent("BeginBatchEdit", ref paramsArray);

			cancel = (Int16)paramsArray[0];
		}

		public void UndoBatchEdit([In] [Out] ref object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("UndoBatchEdit");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			_eventBinding.RaiseCustomEvent("UndoBatchEdit", ref paramsArray);

			cancel = (Int16)paramsArray[0];
		}

		public void BeforeBeginTransaction([In] [Out] ref object cancel, [In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeBeginTransaction");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel, connection);
				return;
			}

			object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(connection, 1);
			_eventBinding.RaiseCustomEvent("BeforeBeginTransaction", ref paramsArray);

			cancel = (Int16)paramsArray[0];
			connection = (NetOffice.ADODBApi.Connection)paramsArray[1];
		}

		public void AfterBeginTransaction([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterBeginTransaction");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(connection);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(connection, 0);
			_eventBinding.RaiseCustomEvent("AfterBeginTransaction", ref paramsArray);

			connection = (NetOffice.ADODBApi.Connection)paramsArray[0];
		}

		public void BeforeCommitTransaction([In] [Out] ref object cancel, [In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeCommitTransaction");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(cancel, connection);
				return;
			}

			object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(connection, 1);
			_eventBinding.RaiseCustomEvent("BeforeCommitTransaction", ref paramsArray);

			cancel = (Int16)paramsArray[0];
			connection = (NetOffice.ADODBApi.Connection)paramsArray[1];
		}

		public void AfterCommitTransaction([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterCommitTransaction");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(connection);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(connection, 0);
			_eventBinding.RaiseCustomEvent("AfterCommitTransaction", ref paramsArray);

			connection = (NetOffice.ADODBApi.Connection)paramsArray[0];
		}

		public void RollbackTransaction([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object connection)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("RollbackTransaction");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(connection);
				return;
			}

			object[] paramsArray = new object[1];
			paramsArray.SetValue(connection, 0);
			_eventBinding.RaiseCustomEvent("RollbackTransaction", ref paramsArray);

			connection = (NetOffice.ADODBApi.Connection)paramsArray[0];
		}

		public void OnConnect()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("OnConnect");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("OnConnect", ref paramsArray);
		}

		public void OnDisconnect()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("OnDisconnect");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("OnDisconnect", ref paramsArray);
		}

		public void PivotTableChange([In] object reason)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("PivotTableChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reason);
				return;
			}

			Int32 newReason = Convert.ToInt32(reason);
			object[] paramsArray = new object[1];
			paramsArray[0] = newReason;
			_eventBinding.RaiseCustomEvent("PivotTableChange", ref paramsArray);
		}

		public void Query()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("Query");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("Query", ref paramsArray);
		}

		public void BeforeQuery()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeQuery");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("BeforeQuery", ref paramsArray);
		}

		public void SelectionChange()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("SelectionChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("SelectionChange", ref paramsArray);
		}

		public void CommandBeforeExecute([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("CommandBeforeExecute");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(command, cancel);
				return;
			}

			object newCommand = (object)command;
			object newCancel = Factory.CreateObjectFromComProxy(_eventClass, cancel) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newCancel;
			_eventBinding.RaiseCustomEvent("CommandBeforeExecute", ref paramsArray);
		}

		public void CommandChecked([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object _checked)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("CommandChecked");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(command, _checked);
				return;
			}

			object newCommand = (object)command;
			object newChecked = Factory.CreateObjectFromComProxy(_eventClass, _checked) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newChecked;
			_eventBinding.RaiseCustomEvent("CommandChecked", ref paramsArray);
		}

		public void CommandEnabled([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object enabled)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("CommandEnabled");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(command, enabled);
				return;
			}

			object newCommand = (object)command;
			object newEnabled = Factory.CreateObjectFromComProxy(_eventClass, enabled) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newEnabled;
			_eventBinding.RaiseCustomEvent("CommandEnabled", ref paramsArray);
		}

		public void CommandExecute([In] object command)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("CommandExecute");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(command);
				return;
			}

			object newCommand = (object)command;
			object[] paramsArray = new object[1];
			paramsArray[0] = newCommand;
			_eventBinding.RaiseCustomEvent("CommandExecute", ref paramsArray);
		}

		public void DataSetChange()
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DataSetChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray();
				return;
			}

			object[] paramsArray = new object[0];
			_eventBinding.RaiseCustomEvent("DataSetChange", ref paramsArray);
		}

		public void BeforeScreenTip([In, MarshalAs(UnmanagedType.IDispatch)] object screenTipText, [In, MarshalAs(UnmanagedType.IDispatch)] object sourceObject)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeScreenTip");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(screenTipText, sourceObject);
				return;
			}

			object newScreenTipText = Factory.CreateObjectFromComProxy(_eventClass, screenTipText) as object;
			object newSourceObject = Factory.CreateObjectFromComProxy(_eventClass, sourceObject) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newScreenTipText;
			paramsArray[1] = newSourceObject;
			_eventBinding.RaiseCustomEvent("BeforeScreenTip", ref paramsArray);
		}

		public void BeforeRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("BeforeRender");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(drawObject, chartObject, cancel);
				return;
			}

			object newdrawObject = Factory.CreateObjectFromComProxy(_eventClass, drawObject) as object;
			object newchartObject = Factory.CreateObjectFromComProxy(_eventClass, chartObject) as object;
			object newCancel = Factory.CreateObjectFromComProxy(_eventClass, cancel) as object;
			object[] paramsArray = new object[3];
			paramsArray[0] = newdrawObject;
			paramsArray[1] = newchartObject;
			paramsArray[2] = newCancel;
			_eventBinding.RaiseCustomEvent("BeforeRender", ref paramsArray);
		}

		public void AfterRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterRender");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(drawObject, chartObject);
				return;
			}

			object newdrawObject = Factory.CreateObjectFromComProxy(_eventClass, drawObject) as object;
			object newchartObject = Factory.CreateObjectFromComProxy(_eventClass, chartObject) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newdrawObject;
			paramsArray[1] = newchartObject;
			_eventBinding.RaiseCustomEvent("AfterRender", ref paramsArray);
		}

		public void AfterFinalRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterFinalRender");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(drawObject);
				return;
			}

			object newdrawObject = Factory.CreateObjectFromComProxy(_eventClass, drawObject) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdrawObject;
			_eventBinding.RaiseCustomEvent("AfterFinalRender", ref paramsArray);
		}

		public void AfterLayout([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("AfterLayout");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(drawObject);
				return;
			}

			object newdrawObject = Factory.CreateObjectFromComProxy(_eventClass, drawObject) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdrawObject;
			_eventBinding.RaiseCustomEvent("AfterLayout", ref paramsArray);
		}

		public void MouseWheel([In] object page, [In] object count)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("MouseWheel");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(page, count);
				return;
			}

			bool newPage = Convert.ToBoolean(page);
			Int32 newCount = Convert.ToInt32(count);
			object[] paramsArray = new object[2];
			paramsArray[0] = newPage;
			paramsArray[1] = newCount;
			_eventBinding.RaiseCustomEvent("MouseWheel", ref paramsArray);
		}

		public void ViewChange([In] object reason)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("ViewChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reason);
				return;
			}

			Int32 newReason = Convert.ToInt32(reason);
			object[] paramsArray = new object[1];
			paramsArray[0] = newReason;
			_eventBinding.RaiseCustomEvent("ViewChange", ref paramsArray);
		}

		public void DataChange([In] object reason)
		{
			Delegate[] recipients = _eventBinding.GetEventRecipients("DataChange");
			if( (true == _eventClass.IsCurrentlyDisposing) || (recipients.Length == 0) )
			{
				Invoker.ReleaseParamsArray(reason);
				return;
			}

			Int32 newReason = Convert.ToInt32(reason);
			object[] paramsArray = new object[1];
			paramsArray[0] = newReason;
			_eventBinding.RaiseCustomEvent("DataChange", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}