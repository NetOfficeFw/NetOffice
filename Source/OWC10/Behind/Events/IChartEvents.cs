using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OWC10Api.EventContracts.IChartEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class IChartEvents_SinkHelper : SinkHelper, NetOffice.OWC10Api.EventContracts.IChartEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from IChartEvents
        /// </summary>
        public static readonly string Id = "F5B39A7A-1480-11D3-8549-00C04FAC67D7";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public IChartEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region IChartEvents

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
        public void DblClick()
        {
            if (!Validate("DblClick"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("DblClick", ref paramsArray);
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
        /// <param name="keyCode"></param>
        /// <param name="shift"></param>
        public void KeyDown([In] object keyCode, [In] object shift)
        {
            if (!Validate("KeyDown"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift);
                return;
            }

            Int32 newKeyCode = ToInt32(keyCode);
            Int32 newShift = ToInt32(shift);
            object[] paramsArray = new object[2];
            paramsArray[0] = newKeyCode;
            paramsArray[1] = newShift;
            EventBinding.RaiseCustomEvent("KeyDown", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="keyCode"></param>
        /// <param name="shift"></param>
        public void KeyUp([In] object keyCode, [In] object shift)
        {
            if (!Validate("KeyUp"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift);
                return;
            }

            Int32 newKeyCode = ToInt32(keyCode);
            Int32 newShift = ToInt32(shift);
            object[] paramsArray = new object[2];
            paramsArray[0] = newKeyCode;
            paramsArray[1] = newShift;
            EventBinding.RaiseCustomEvent("KeyUp", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="keyAscii"></param>
        public void KeyPress([In] object keyAscii)
        {
            if (!Validate("KeyPress"))
            {
                Invoker.ReleaseParamsArray(keyAscii);
                return;
            }

            Int32 newKeyAscii = ToInt32(keyAscii);
            object[] paramsArray = new object[1];
            paramsArray[0] = newKeyAscii;
            EventBinding.RaiseCustomEvent("KeyPress", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="keyCode"></param>
        /// <param name="shift"></param>
        /// <param name="cancel"></param>
        public void BeforeKeyDown([In] object keyCode, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
        {
            if (!Validate("BeforeKeyDown"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift, cancel);
                return;
            }

            Int32 newKeyCode = ToInt32(keyCode);
            Int32 newShift = ToInt32(shift);
            NetOffice.OWC10Api.ByRef newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, cancel, typeof(NetOffice.OWC10Api.ByRef));
            object[] paramsArray = new object[3];
            paramsArray[0] = newKeyCode;
            paramsArray[1] = newShift;
            paramsArray[2] = newCancel;
            EventBinding.RaiseCustomEvent("BeforeKeyDown", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="keyCode"></param>
        /// <param name="shift"></param>
        /// <param name="cancel"></param>
        public void BeforeKeyUp([In] object keyCode, [In] object shift, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
        {
            if (!Validate("BeforeKeyUp"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift, cancel);
                return;
            }

            Int32 newKeyCode = Convert.ToInt32(keyCode);
            Int32 newShift = Convert.ToInt32(shift);
            NetOffice.OWC10Api.ByRef newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, cancel, typeof(NetOffice.OWC10Api.ByRef));
            object[] paramsArray = new object[3];
            paramsArray[0] = newKeyCode;
            paramsArray[1] = newShift;
            paramsArray[2] = newCancel;
            EventBinding.RaiseCustomEvent("BeforeKeyUp", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="keyAscii"></param>
        /// <param name="cancel"></param>
        public void BeforeKeyPress([In] object keyAscii, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
        {
            if (!Validate("BeforeKeyPress"))
            {
                Invoker.ReleaseParamsArray(keyAscii, cancel);
                return;
            }

            Int32 newKeyAscii = Convert.ToInt32(keyAscii);
            NetOffice.OWC10Api.ByRef newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, cancel, typeof(NetOffice.OWC10Api.ByRef));
            object[] paramsArray = new object[2];
            paramsArray[0] = newKeyAscii;
            paramsArray[1] = newCancel;
            EventBinding.RaiseCustomEvent("BeforeKeyPress", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        public void MouseDown([In] object button, [In] object shift, [In] object x, [In] object y)
        {
            if (!Validate("MouseDown"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }

            Int32 newButton = ToInt32(button);
            Int32 newShift = ToInt32(shift);
            Int32 newx = ToInt32(x);
            Int32 newy = ToInt32(y);
            object[] paramsArray = new object[4];
            paramsArray[0] = newButton;
            paramsArray[1] = newShift;
            paramsArray[2] = newx;
            paramsArray[3] = newy;
            EventBinding.RaiseCustomEvent("MouseDown", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        public void MouseMove([In] object button, [In] object shift, [In] object x, [In] object y)
        {
            if (!Validate("MouseMove"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }

            Int32 newButton = ToInt32(button);
            Int32 newShift = ToInt32(shift);
            Int32 newx = ToInt32(x);
            Int32 newy = ToInt32(y);
            object[] paramsArray = new object[4];
            paramsArray[0] = newButton;
            paramsArray[1] = newShift;
            paramsArray[2] = newx;
            paramsArray[3] = newy;
            EventBinding.RaiseCustomEvent("MouseMove", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        public void MouseUp([In] object button, [In] object shift, [In] object x, [In] object y)
        {
            if (!Validate("MouseUp"))
            {
                Invoker.ReleaseParamsArray(button, shift, x, y);
                return;
            }

            Int32 newButton = ToInt32(button);
            Int32 newShift = ToInt32(shift);
            Int32 newx = ToInt32(x);
            Int32 newy = ToInt32(y);
            object[] paramsArray = new object[4];
            paramsArray[0] = newButton;
            paramsArray[1] = newShift;
            paramsArray[2] = newx;
            paramsArray[3] = newy;
            EventBinding.RaiseCustomEvent("MouseUp", ref paramsArray);
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
        /// <param name="tipText"></param>
        /// <param name="contextObject"></param>
        public void BeforeScreenTip([In, MarshalAs(UnmanagedType.IDispatch)] object tipText, [In, MarshalAs(UnmanagedType.IDispatch)] object contextObject)
        {
            if (!Validate("BeforeScreenTip"))
            {
                Invoker.ReleaseParamsArray(tipText, contextObject);
                return;
            }

            NetOffice.OWC10Api.ByRef newTipText = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, tipText, typeof(NetOffice.OWC10Api.ByRef));
            object newContextObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, contextObject) as object;
            object[] paramsArray = new object[2];
            paramsArray[0] = newTipText;
            paramsArray[1] = newContextObject;
            EventBinding.RaiseCustomEvent("BeforeScreenTip", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="command"></param>
        /// <param name="enabled"></param>
        public void CommandEnabled([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object enabled)
        {
            if (!Validate("CommandEnabled"))
            {
                Invoker.ReleaseParamsArray(command, enabled);
                return;
            }

            object newCommand = (object)command;
            NetOffice.OWC10Api.ByRef newEnabled = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, enabled, typeof(NetOffice.OWC10Api.ByRef));
            object[] paramsArray = new object[2];
            paramsArray[0] = newCommand;
            paramsArray[1] = newEnabled;
            EventBinding.RaiseCustomEvent("CommandEnabled", ref paramsArray);
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
            NetOffice.OWC10Api.ByRef newChecked = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, _checked, typeof(NetOffice.OWC10Api.ByRef));
            object[] paramsArray = new object[2];
            paramsArray[0] = newCommand;
            paramsArray[1] = newChecked;
            EventBinding.RaiseCustomEvent("CommandChecked", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="command"></param>
        /// <param name="caption"></param>
        public void CommandTipText([In] object command, [In, MarshalAs(UnmanagedType.IDispatch)] object caption)
        {
            if (!Validate("CommandTipText"))
            {
                Invoker.ReleaseParamsArray(command, caption);
                return;
            }

            object newCommand = (object)command;
            NetOffice.OWC10Api.ByRef newCaption = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, command, typeof(NetOffice.OWC10Api.ByRef));
            object[] paramsArray = new object[2];
            paramsArray[0] = newCommand;
            paramsArray[1] = newCaption;
            EventBinding.RaiseCustomEvent("CommandTipText", ref paramsArray);
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
            NetOffice.OWC10Api.ByRef newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, cancel, typeof(NetOffice.OWC10Api.ByRef));
            object[] paramsArray = new object[2];
            paramsArray[0] = newCommand;
            paramsArray[1] = newCancel;
            EventBinding.RaiseCustomEvent("CommandBeforeExecute", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="command"></param>
        /// <param name="succeeded"></param>
        public void CommandExecute([In] object command, [In] object succeeded)
        {
            if (!Validate("CommandExecute"))
            {
                Invoker.ReleaseParamsArray(command, succeeded);
                return;
            }

            object newCommand = (object)command;
            bool newSucceeded = ToBoolean(succeeded);
            object[] paramsArray = new object[2];
            paramsArray[0] = newCommand;
            paramsArray[1] = newSucceeded;
            EventBinding.RaiseCustomEvent("CommandExecute", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="menu"></param>
        /// <param name="cancel"></param>
        public void BeforeContextMenu([In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object menu, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
        {
            if (!Validate("BeforeContextMenu"))
            {
                Invoker.ReleaseParamsArray(x, y, menu, cancel);
                return;
            }

            Int32 newx = ToInt32(x);
            Int32 newy = ToInt32(y);
            NetOffice.OWC10Api.ByRef newMenu = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, menu, typeof(NetOffice.OWC10Api.ByRef));
            NetOffice.OWC10Api.ByRef newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, cancel, typeof(NetOffice.OWC10Api.ByRef));
            object[] paramsArray = new object[4];
            paramsArray[0] = newx;
            paramsArray[1] = newy;
            paramsArray[2] = newMenu;
            paramsArray[3] = newCancel;
            EventBinding.RaiseCustomEvent("BeforeContextMenu", ref paramsArray);
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

            NetOffice.OWC10Api.ChChartDraw newdrawObject = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ChChartDraw>(EventClass, drawObject, typeof(NetOffice.OWC10Api.ChChartDraw));
            object newchartObject = Factory.CreateEventArgumentObjectFromComProxy(EventClass, chartObject) as object;
            NetOffice.OWC10Api.ByRef newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ByRef>(EventClass, cancel, typeof(NetOffice.OWC10Api.ByRef));
            object[] paramsArray = new object[3];
            paramsArray[0] = newdrawObject;
            paramsArray[1] = newchartObject;
            paramsArray[2] = newCancel;
            EventBinding.RaiseCustomEvent("BeforeRender", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="drawObject"></param>
        /// <param name="chartObject"></param>
        public void AfterRender([In, MarshalAs(UnmanagedType.IDispatch)] object drawObject, [In, MarshalAs(UnmanagedType.IDispatch)] object chartObject)
        {
            if (!Validate("AfterRender"))
            {
                Invoker.ReleaseParamsArray(drawObject, chartObject);
                return;
            }

            NetOffice.OWC10Api.ChChartDraw newdrawObject = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ChChartDraw>(EventClass, drawObject, typeof(NetOffice.OWC10Api.ChChartDraw));
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

            NetOffice.OWC10Api.ChChartDraw newdrawObject = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ChChartDraw>(EventClass, drawObject, typeof(NetOffice.OWC10Api.ChChartDraw));
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

            NetOffice.OWC10Api.ChChartDraw newdrawObject = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.ChChartDraw>(EventClass, drawObject, typeof(NetOffice.OWC10Api.ChChartDraw));
            object[] paramsArray = new object[1];
            paramsArray[0] = newdrawObject;
            EventBinding.RaiseCustomEvent("AfterLayout", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        public void ViewChange()
        {
            if (!Validate("ViewChange"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("ViewChange", ref paramsArray);
        }

        #endregion
    }
}
