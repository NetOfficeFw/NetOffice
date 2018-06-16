using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.MSComctlLibApi.EventContracts.DImageComboEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class DImageComboEvents_SinkHelper : SinkHelper, NetOffice.MSComctlLibApi.EventContracts.DImageComboEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from DImageComboEvents
        /// </summary>
        public static readonly string Id = "DD9DA665-8594-11D1-B16A-00C0F0283628";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public DImageComboEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region DImageComboEvents

        /// <summary>
        /// 
        /// </summary>
        public void Change()
        {
            if (!Validate("Change"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Change", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        public void Dropdown()
        {
            if (!Validate("Dropdown"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Dropdown", ref paramsArray);
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

            Int16 newKeyCode = ToInt16(keyCode);
            Int16 newShift = ToInt16(shift);
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

            Int16 newKeyCode = ToInt16(keyCode);
            Int16 newShift = ToInt16(shift);
            object[] paramsArray = new object[2];
            paramsArray[0] = newKeyCode;
            paramsArray[1] = newShift;
            EventBinding.RaiseCustomEvent("KeyUp", ref paramsArray);
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
        /// <param name="data"></param>
        /// <param name="allowedEffects"></param>
        public void OLEStartDrag([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object data, [In] [Out] ref object allowedEffects)
        {
            if (!Validate("OLEStartDrag"))
            {
                Invoker.ReleaseParamsArray(data, allowedEffects);
                return;
            }

            NetOffice.MSComctlLibApi.DataObject newData = Factory.CreateKnownObjectFromComProxy<NetOffice.MSComctlLibApi.DataObject>(EventClass, data, typeof(NetOffice.MSComctlLibApi.DataObject));
            (newData as ICOMProxyShareProvider).GetProxyShare().Acquire();

            object[] paramsArray = new object[2];
            paramsArray.SetValue(newData, 0);
            paramsArray.SetValue(allowedEffects, 1);
            EventBinding.RaiseCustomEvent("OLEStartDrag", ref paramsArray);

            data = newData.UnderlyingObject;
            allowedEffects = ToInt32(paramsArray[1]);

            (newData as ICOMProxyShareProvider).GetProxyShare().Release();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="effect"></param>
        /// <param name="defaultCursors"></param>
        public void OLEGiveFeedback([In] [Out] ref object effect, [In] [Out] ref object defaultCursors)
        {
            if (!Validate("OLEGiveFeedback"))
            {
                Invoker.ReleaseParamsArray(effect, defaultCursors);
                return;
            }

            object[] paramsArray = new object[2];
            paramsArray.SetValue(effect, 0);
            paramsArray.SetValue(defaultCursors, 1);
            EventBinding.RaiseCustomEvent("OLEGiveFeedback", ref paramsArray);

            effect = ToInt32(paramsArray[0]);
            defaultCursors = ToBoolean(paramsArray[1]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <param name="dataFormat"></param>
        public void OLESetData([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object data, [In] [Out] ref object dataFormat)
        {
            if (!Validate("OLESetData"))
            {
                Invoker.ReleaseParamsArray(data, dataFormat);
                return;
            }

            NetOffice.MSComctlLibApi.DataObject newData = Factory.CreateKnownObjectFromComProxy<NetOffice.MSComctlLibApi.DataObject>(EventClass, data, typeof(NetOffice.MSComctlLibApi.DataObject));
            (newData as ICOMProxyShareProvider).GetProxyShare().Acquire();

            object[] paramsArray = new object[2];
            paramsArray.SetValue(newData, 0);
            paramsArray.SetValue(dataFormat, 1);
            EventBinding.RaiseCustomEvent("OLESetData", ref paramsArray);

            data = newData.UnderlyingObject;
            dataFormat = ToInt16(paramsArray[1]);

            (newData as ICOMProxyShareProvider).GetProxyShare().Release();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="effect"></param>
        public void OLECompleteDrag([In] [Out] ref object effect)
        {
            if (!Validate("OLECompleteDrag"))
            {
                Invoker.ReleaseParamsArray(effect);
                return;
            }

            object[] paramsArray = new object[1];
            paramsArray.SetValue(effect, 0);
            EventBinding.RaiseCustomEvent("OLECompleteDrag", ref paramsArray);

            effect = ToInt32(paramsArray[0]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <param name="effect"></param>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="state"></param>
        public void OLEDragOver([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object data, [In] [Out] ref object effect, [In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y, [In] [Out] ref object state)
        {
            if (!Validate("OLEDragOver"))
            {
                Invoker.ReleaseParamsArray(data, effect, button, shift, x, y, state);
                return;
            }

            NetOffice.MSComctlLibApi.DataObject newData = Factory.CreateKnownObjectFromComProxy<NetOffice.MSComctlLibApi.DataObject>(EventClass, data, typeof(NetOffice.MSComctlLibApi.DataObject));
            (newData as ICOMProxyShareProvider).GetProxyShare().Acquire();

            object[] paramsArray = new object[7];
            paramsArray.SetValue(newData, 0);
            paramsArray.SetValue(effect, 1);
            paramsArray.SetValue(button, 2);
            paramsArray.SetValue(shift, 3);
            paramsArray.SetValue(x, 4);
            paramsArray.SetValue(y, 5);
            paramsArray.SetValue(state, 6);
            EventBinding.RaiseCustomEvent("OLEDragOver", ref paramsArray);

            data = newData.UnderlyingObject;
            effect = ToInt32(paramsArray[1]);
            button = ToInt16(paramsArray[2]);
            shift = ToInt16(paramsArray[3]);
            x = ToSingle(paramsArray[4]);
            y = ToSingle(paramsArray[5]);
            state = ToInt16(paramsArray[6]);

            (newData as ICOMProxyShareProvider).GetProxyShare().Release();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <param name="effect"></param>
        /// <param name="button"></param>
        /// <param name="shift"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        public void OLEDragDrop([In] [Out, MarshalAs(UnmanagedType.IDispatch)] object data, [In] [Out] ref object effect, [In] [Out] ref object button, [In] [Out] ref object shift, [In] [Out] ref object x, [In] [Out] ref object y)
        {
            if (!Validate("OLEDragDrop"))
            {
                Invoker.ReleaseParamsArray(data, effect, button, shift, x, y);
                return;
            }

            NetOffice.MSComctlLibApi.DataObject newData = Factory.CreateKnownObjectFromComProxy<NetOffice.MSComctlLibApi.DataObject>(EventClass, data, typeof(NetOffice.MSComctlLibApi.DataObject));
            (newData as ICOMProxyShareProvider).GetProxyShare().Acquire();

            object[] paramsArray = new object[6];
            paramsArray.SetValue(newData, 0);
            paramsArray.SetValue(effect, 1);
            paramsArray.SetValue(button, 2);
            paramsArray.SetValue(shift, 3);
            paramsArray.SetValue(x, 4);
            paramsArray.SetValue(y, 5);
            EventBinding.RaiseCustomEvent("OLEDragDrop", ref paramsArray);

            data = newData.UnderlyingObject;
            effect = ToInt32(paramsArray[1]);
            button = ToInt16(paramsArray[2]);
            shift = ToInt16(paramsArray[3]);
            x = ToSingle(paramsArray[4]);
            y = ToSingle(paramsArray[5]);

            (newData as ICOMProxyShareProvider).GetProxyShare().Release();
        }

        #endregion
    }
}
