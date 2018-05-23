using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice.Exceptions;
using NetOffice.Attributes;
using NetOffice.VBIDEApi.EventInterfaces;

namespace NetOffice.VBIDEApi.Behind.EventInterfaces
{
    /// <summary>
    /// Default implementation of <see cref="_dispCommandBarControlEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class _dispCommandBarControlEvents_SinkHelper : SinkHelper, _dispCommandBarControlEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from _dispCommandBarControlEvents
        /// </summary>
        public static readonly string Id = "0002E131-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        /// <exception cref="NetOfficeCOMException">Unexpected error</exception>
        public _dispCommandBarControlEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region _dispCommandBarControlEvents

        /// <summary>
        /// 
        /// </summary>
        /// <param name="commandBarControl"></param>
        /// <param name="handled"></param>
        /// <param name="cancelDefault"></param>
        public void Click([In, MarshalAs(UnmanagedType.IDispatch)] object commandBarControl, [In] [Out] ref object handled, [In] [Out] ref object cancelDefault)
        {
            if (!Validate("Click"))
            {
                Invoker.ReleaseParamsArray(commandBarControl, handled, cancelDefault);
                return;
            }

            object newCommandBarControl = Factory.CreateEventArgumentObjectFromComProxy(EventClass, commandBarControl) as object;
            object[] paramsArray = new object[3];
            paramsArray[0] = newCommandBarControl;
            paramsArray.SetValue(handled, 1);
            paramsArray.SetValue(cancelDefault, 2);
            EventBinding.RaiseCustomEvent("Click", ref paramsArray);

            handled = ToBoolean(paramsArray[1]);
            cancelDefault = ToBoolean(paramsArray[2]);
        }

        #endregion
    }
}
