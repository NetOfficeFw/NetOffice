using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Exceptions;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OfficeApi.EventContracts._CustomTaskPaneEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class _CustomTaskPaneEvents_SinkHelper : SinkHelper, NetOffice.OfficeApi.EventContracts._CustomTaskPaneEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from _CustomTaskPaneEvents
        /// </summary>
        public static readonly string Id = "000C033C-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        /// <exception cref="NetOfficeCOMException">Unexpected error</exception>
        public _CustomTaskPaneEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region _CustomTaskPaneEvents

        /// <summary>
        /// 
        /// </summary>
        /// <param name="customTaskPaneInst"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="customTaskPaneInst"></param>
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
