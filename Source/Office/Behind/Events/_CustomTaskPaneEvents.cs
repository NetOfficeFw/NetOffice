using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.OfficeApi.EventInterfaces;

namespace NetOffice.OfficeApi.Behind.EventInterfaces
{
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class _CustomTaskPaneEvents_SinkHelper : SinkHelper, NetOffice.OfficeApi.EventInterfaces._CustomTaskPaneEvents
    {
        #region Static

        public static readonly string Id = "000C033C-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        public _CustomTaskPaneEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region _CustomTaskPaneEvents

        public void VisibleStateChange([In, MarshalAs(UnmanagedType.IDispatch)] object customTaskPaneInst)
        {
            if (!Validate("VisibleStateChange"))
            {
                Invoker.ReleaseParamsArray(customTaskPaneInst);
                return;
            }

            NetOffice.OfficeApi._CustomTaskPane newCustomTaskPaneInst = Factory.CreateEventArgumentObjectFromComProxy(EventClass, customTaskPaneInst) as NetOffice.OfficeApi._CustomTaskPane;
            object[] paramsArray = new object[1];
            paramsArray[0] = newCustomTaskPaneInst;
            EventBinding.RaiseCustomEvent("VisibleStateChange", ref paramsArray);
        }

        public void DockPositionStateChange([In, MarshalAs(UnmanagedType.IDispatch)] object customTaskPaneInst)
        {
            if (!Validate("DockPositionStateChange"))
            {
                Invoker.ReleaseParamsArray(customTaskPaneInst);
                return;
            }

            NetOffice.OfficeApi._CustomTaskPane newCustomTaskPaneInst = Factory.CreateEventArgumentObjectFromComProxy(EventClass, customTaskPaneInst) as NetOffice.OfficeApi._CustomTaskPane;
            object[] paramsArray = new object[1];
            paramsArray[0] = newCustomTaskPaneInst;
            EventBinding.RaiseCustomEvent("DockPositionStateChange", ref paramsArray);
        }

        #endregion
    }
}
