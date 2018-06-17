using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice.AccessApi.Behind.EventContracts
{

	/// <summary>
	/// Default implementation of <see cref="NetOffice.AccessApi.EventContracts._FormEvents2"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _FormEvents2_SinkHelper : SinkHelper, NetOffice.AccessApi.EventContracts._FormEvents2
	{
		#region Static
		
		/// <summary>
		/// Interface Id from _FormEvents2
		/// </summary>
		public static readonly string Id = "0EA530DD-5B30-4278-BD28-47C4D11619BD";
		
		#endregion
		
		#region Ctor

		/// <summary>
		/// Creates an instance of the class
		/// </summary>
		/// <param name="eventClass"></param>
		/// <param name="connectPoint"></param>
		/// <exception cref="NetOfficeCOMException">Unexpected error</exception>
		public _FormEvents2_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion		

		#region _FormEvents2
		
		/// <summary>
		/// 
		/// </summary>
		public void Load()
        {
            if (!Validate("Load"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Load", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void Current()
        {
            if (!Validate("Current"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Current", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		public void BeforeInsert([In] [Out] ref object cancel)
        {
            if (!Validate("BeforeInsert"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeInsert", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
		}

		/// <summary>
		/// 
		/// </summary>
		public void AfterInsert()
        {
            if (!Validate("AfterInsert"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("AfterInsert", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		public void BeforeUpdate([In] [Out] ref object cancel)
		{
            if (!Validate("BeforeUpdate"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeforeUpdate", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
		}

		/// <summary>
		/// 
		/// </summary>
		public void AfterUpdate()
        {
            if (!Validate("AfterUpdate"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("AfterUpdate", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		public void Delete([In] [Out] ref object cancel)
		{
            if (!Validate("Delete"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("Delete", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		/// <param name="response"></param>
		public void BeforeDelConfirm([In] [Out] ref object cancel, [In] [Out] ref object response)
        {
            if (!Validate("BeforeDelConfirm"))
            {
                Invoker.ReleaseParamsArray(cancel, response);
                return;
            }

			object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(response, 1);
			EventBinding.RaiseCustomEvent("BeforeDelConfirm", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
            response = ToInt16(paramsArray[1]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="status"></param>
		public void AfterDelConfirm([In] [Out] ref object status)
        {
            if (!Validate("AfterDelConfirm"))
            {
                Invoker.ReleaseParamsArray(status);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(status, 0);
			EventBinding.RaiseCustomEvent("AfterDelConfirm", ref paramsArray);

			status = ToInt16(paramsArray[0]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		public void Open([In] [Out] ref object cancel)
        {
            if (!Validate("Open"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("Open", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
        }

		/// <summary>
		/// 
		/// </summary>
		public void Resize()
        {
            if (!Validate("Resize"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Resize", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		public void Unload([In] [Out] ref object cancel)
        {
            if (!Validate("Unload"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("Unload", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
        }

		/// <summary>
		/// 
		/// </summary>
		public void Close()
		{
            if (!Validate("Close"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Close", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void Activate()
        {
            if (!Validate("Activate"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Activate", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void Deactivate()
		{
            if (!Validate("Deactivate"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Deactivate", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void GotFocus()
        {
            if (!Validate("GotFocus"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("GotFocus", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void LostFocus()
		{
            if (!Validate("LostFocus"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("LostFocus", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void Click()
        {
            if (!Validate("Click"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Click", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		public void DblClick([In] [Out] ref object cancel)
		{
            if (!Validate("DblClick"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("DblClick", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="button"></param>
		/// <param name="shift"></param>
		/// <param name="x"></param>
		/// <param name="y"></param>
		public void MouseDown([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y)
		{
            if (!Validate("MouseDown"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }

			object[] paramsArray = new object[4];
			paramsArray.SetValue(button, 0);
			paramsArray.SetValue(shift, 1);
			paramsArray.SetValue(x, 2);
			paramsArray.SetValue(y, 3);
			EventBinding.RaiseCustomEvent("MouseDown", ref paramsArray);

			button = ToInt16(paramsArray[0]);
			shift = ToInt16(paramsArray[1]);
			x = ToSingle(paramsArray[2]);
			y = ToSingle(paramsArray[3]);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="button"></param>
		/// <param name="shift"></param>
		/// <param name="x"></param>
		/// <param name="y"></param>
		public void MouseMove([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y)
		{
            if (!Validate("MouseMove"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }

			object[] paramsArray = new object[4];
			paramsArray.SetValue(button, 0);
			paramsArray.SetValue(shift, 1);
			paramsArray.SetValue(x, 2);
			paramsArray.SetValue(y, 3);
			EventBinding.RaiseCustomEvent("MouseMove", ref paramsArray);

            button = ToInt16(paramsArray[0]);
            shift = ToInt16(paramsArray[1]);
            x = ToSingle(paramsArray[2]);
            y = ToSingle(paramsArray[3]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="button"></param>
		/// <param name="shift"></param>
		/// <param name="x"></param>
		/// <param name="y"></param>
		public void MouseUp([In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y)
		{
            if (!Validate("MouseUp"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }

			object[] paramsArray = new object[4];
			paramsArray.SetValue(button, 0);
			paramsArray.SetValue(shift, 1);
			paramsArray.SetValue(x, 2);
			paramsArray.SetValue(y, 3);
			EventBinding.RaiseCustomEvent("MouseUp", ref paramsArray);

            button = ToInt16(paramsArray[0]);
            shift = ToInt16(paramsArray[1]);
            x = ToSingle(paramsArray[2]);
            y = ToSingle(paramsArray[3]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="keyCode"></param>
		/// <param name="shift"></param>
		public void KeyDown([In] [Out] ref object keyCode, [In] [Out] ref object shift)
		{
            if (!Validate("KeyDown"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift);
                return;
            }

			object[] paramsArray = new object[2];
			paramsArray.SetValue(keyCode, 0);
			paramsArray.SetValue(shift, 1);
			EventBinding.RaiseCustomEvent("KeyDown", ref paramsArray);

			keyCode = ToInt16(paramsArray[0]);
            shift = ToInt16(paramsArray[1]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="keyAscii"></param>
		public void KeyPress([In] [Out] ref object keyAscii)
		{
            if (!Validate("KeyPress"))
            {
                Invoker.ReleaseParamsArray(keyAscii);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(keyAscii, 0);
			EventBinding.RaiseCustomEvent("KeyPress", ref paramsArray);

			keyAscii = ToInt16(paramsArray[0]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="keyCode"></param>
		/// <param name="shift"></param>
		public void KeyUp([In] [Out] ref object keyCode, [In] [Out] ref object shift)
        {
            if (!Validate("KeyUp"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift);
                return;
            }

			object[] paramsArray = new object[2];
			paramsArray.SetValue(keyCode, 0);
			paramsArray.SetValue(shift, 1);
			EventBinding.RaiseCustomEvent("KeyUp", ref paramsArray);

			keyCode = ToInt16(paramsArray[0]);
            shift = ToInt16(paramsArray[1]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="dataErr"></param>
		/// <param name="response"></param>
		public void Error([In] [Out] ref object dataErr, [In] [Out] ref object response)
        {
            if (!Validate("Error"))
            {
                Invoker.ReleaseParamsArray(dataErr, response);
                return;
            }

			object[] paramsArray = new object[2];
			paramsArray.SetValue(dataErr, 0);
			paramsArray.SetValue(response, 1);
			EventBinding.RaiseCustomEvent("Error", ref paramsArray);

			dataErr = ToInt16(paramsArray[0]);
            response = ToInt16(paramsArray[1]);
        }

		/// <summary>
		/// 
		/// </summary>
		public void Timer()
		{
            if (!Validate("Timer"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Timer", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		/// <param name="filterType"></param>
		public void Filter([In] [Out] ref object cancel, [In] [Out] ref object filterType)
        {
            if (!Validate("Filter"))
            {
                Invoker.ReleaseParamsArray(cancel, filterType);
                return;
            }

			object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(filterType, 1);
			EventBinding.RaiseCustomEvent("Filter", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
            filterType = ToInt16(paramsArray[1]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		/// <param name="applyType"></param>
		public void ApplyFilter([In] [Out] ref object cancel, [In] [Out] ref object applyType)
        {
            if (!Validate("ApplyFilter"))
            {
                Invoker.ReleaseParamsArray(cancel, applyType);
                return;
            }

			object[] paramsArray = new object[2];
			paramsArray.SetValue(cancel, 0);
			paramsArray.SetValue(applyType, 1);
			EventBinding.RaiseCustomEvent("ApplyFilter", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
            applyType = ToInt16(paramsArray[1]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		public void Dirty([In] [Out] ref object cancel)
		{
            if (!Validate("Dirty"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("Dirty", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		public void Undo([In] [Out] ref object cancel)
        {
            if (!Validate("Undo"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("Undo", ref paramsArray);

			cancel = ToInt16(paramsArray[0]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		public void RecordExit([In] [Out] ref object cancel)
		{
            if (!Validate("RecordExit"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("RecordExit", ref paramsArray);

            cancel = ToInt16(paramsArray[0]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		public void BeginBatchEdit([In] [Out] ref object cancel)
		{
            if (!Validate("BeginBatchEdit"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

			object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("BeginBatchEdit", ref paramsArray);

            cancel = ToInt16(paramsArray[0]);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cancel"></param>
		public void UndoBatchEdit([In] [Out] ref object cancel)
		{
            if (!Validate("UndoBatchEdit"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
			paramsArray.SetValue(cancel, 0);
			EventBinding.RaiseCustomEvent("UndoBatchEdit", ref paramsArray);

            cancel = ToInt16(paramsArray[0]);
        }

		/// <summary>
		/// 
		/// </summary>
		public void OnConnect()
		{
            if (!Validate("OnConnect"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("OnConnect", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void OnDisconnect()
		{
            if (!Validate("OnDisconnect"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("OnDisconnect", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="reason"></param>
		public void PivotTableChange([In] object reason)
        {
            if (!Validate("PivotTableChange"))
            {
                Invoker.ReleaseParamsArray(reason);
                return;
            }

			Int32 newReason = ToInt32(reason);
			object[] paramsArray = new object[1];
			paramsArray[0] = newReason;
			EventBinding.RaiseCustomEvent("PivotTableChange", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void Query()
		{
            if (!Validate("Query"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Query", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void BeforeQuery()
		{
            if (!Validate("BeforeQuery"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("BeforeQuery", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void SelectionChange()
		{
            if (!Validate("SelectionChange"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("SelectionChange", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="command"></param>
        /// <param name="cancel"></param>
        public void CommandBeforeExecute([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
        {
            if (!Validate("CommandBeforeExecute"))
            {
                Invoker.ReleaseParamsArray(command, cancel);
                return;
            }

			object newCommand = (object)command;
			object newCancel = Factory.CreateEventArgumentObjectFromComProxy(EventClass, cancel) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newCancel;
			EventBinding.RaiseCustomEvent("CommandBeforeExecute", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="command"></param>
        /// <param name="_checked"></param>
        public void CommandChecked([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object _checked)
		{
            if (!Validate("CommandChecked"))
            {
                Invoker.ReleaseParamsArray(command, _checked);
                return;
            }

			object newCommand = (object)command;
			object newChecked = Factory.CreateEventArgumentObjectFromComProxy(EventClass, _checked) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newChecked;
			EventBinding.RaiseCustomEvent("CommandChecked", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="command"></param>
        /// <param name="enabled"></param>
        public void CommandEnabled([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object enabled)
        {
            if (!Validate("CommandChecked"))
            {
                Invoker.ReleaseParamsArray(command, enabled);
                return;
            }

			object newCommand = (object)command;
			object newEnabled = Factory.CreateEventArgumentObjectFromComProxy(EventClass, enabled) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newCommand;
			paramsArray[1] = newEnabled;
			EventBinding.RaiseCustomEvent("CommandEnabled", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="command"></param>
		public void CommandExecute([In] object command)
		{
            if (!Validate("CommandExecute"))
            {
                Invoker.ReleaseParamsArray(command);
                return;
            }

            object newCommand = (object)command;
			object[] paramsArray = new object[1];
			paramsArray[0] = newCommand;
			EventBinding.RaiseCustomEvent("CommandExecute", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		public void DataSetChange()
		{
            if (!Validate("DataSetChange"))
            {
                return;
            }

			object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("DataSetChange", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="screenTipText"></param>
        /// <param name="sourceObject"></param>
        public void BeforeScreenTip([In, MarshalAs(UnmanagedType.IDispatch)] object screenTipText, [In, MarshalAs(UnmanagedType.IDispatch)] object sourceObject)
        {
            if (!Validate("BeforeScreenTip"))
            {
                Invoker.ReleaseParamsArray(screenTipText, sourceObject);
                return;
            }

			object newScreenTipText = Factory.CreateEventArgumentObjectFromComProxy(EventClass, screenTipText) as object;
			object newSourceObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sourceObject) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newScreenTipText;
			paramsArray[1] = newSourceObject;
			EventBinding.RaiseCustomEvent("BeforeScreenTip", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="drawObject"></param>
        /// <param name="chartObject"></param>
        /// <param name="cancel"></param>
        public void BeforeRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
        {
            if (!Validate("BeforeRender"))
            {
                Invoker.ReleaseParamsArray(drawObject, chartObject, cancel);
                return;
            }

			object newdrawObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, drawObject) as object;
			object newchartObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, chartObject) as object;
			object newCancel = Factory.CreateEventArgumentObjectFromComProxy(EventClass, cancel) as object;
			object[] paramsArray = new object[3];
			paramsArray[0] = newdrawObject;
			paramsArray[1] = newchartObject;
			paramsArray[2] = newCancel;
			EventBinding.RaiseCustomEvent("BeforeRender", ref paramsArray);
		}

        /// <summary>
        /// 
        /// </summary>
        ///  <param name="drawObject"></param>
        /// <param name="chartObject"></param>
        public void AfterRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject)
		{
            if (!Validate("AfterRender"))
            {
                Invoker.ReleaseParamsArray(drawObject, chartObject);
                return;
            }

			object newdrawObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, drawObject) as object;
			object newchartObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, chartObject) as object;
			object[] paramsArray = new object[2];
			paramsArray[0] = newdrawObject;
			paramsArray[1] = newchartObject;
			EventBinding.RaiseCustomEvent("AfterRender", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="drawObject"></param>
		public void AfterFinalRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject)
        {
            if (!Validate("AfterFinalRender"))
            {
                Invoker.ReleaseParamsArray(drawObject);
                return;
            }

			object newdrawObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, drawObject) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdrawObject;
			EventBinding.RaiseCustomEvent("AfterFinalRender", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="drawObject"></param>
		public void AfterLayout([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject)
        {
            if (!Validate("AfterLayout"))
            {
                Invoker.ReleaseParamsArray(drawObject);
                return;
            }

			object newdrawObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, drawObject) as object;
			object[] paramsArray = new object[1];
			paramsArray[0] = newdrawObject;
			EventBinding.RaiseCustomEvent("AfterLayout", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="page"></param>
		/// <param name="count"></param>
		public void MouseWheel([In] object page, [In] object count)
        {
            if (!Validate("MouseWheel"))
            {
                Invoker.ReleaseParamsArray(page, count);
                return;
            }

			bool newPage = ToBoolean(page);
			Int32 newCount = ToInt32(count);
			object[] paramsArray = new object[2];
			paramsArray[0] = newPage;
			paramsArray[1] = newCount;
			EventBinding.RaiseCustomEvent("MouseWheel", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="reason"></param>
		public void ViewChange([In] object reason)
		{
            if (!Validate("ViewChange"))
            {
                Invoker.ReleaseParamsArray(reason);
                return;
            }

			Int32 newReason = ToInt32(reason);
			object[] paramsArray = new object[1];
			paramsArray[0] = newReason;
			EventBinding.RaiseCustomEvent("ViewChange", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="reason"></param>
		public void DataChange([In] object reason)
		{
            if (!Validate("DataChange"))
            {
                Invoker.ReleaseParamsArray(reason);
                return;
            }

			Int32 newReason = ToInt32(reason);
			object[] paramsArray = new object[1];
			paramsArray[0] = newReason;
			EventBinding.RaiseCustomEvent("DataChange", ref paramsArray);
		}

		#endregion
	}
	
}
