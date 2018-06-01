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
    /// Default implementation of <see cref="NetOffice.OfficeApi.EventContracts._CommandBarButtonEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class _CommandBarButtonEvents_SinkHelper : SinkHelper, NetOffice.OfficeApi.EventContracts._CommandBarButtonEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from _CommandBarButtonEvents
        /// </summary>
        public static readonly string Id = "000C0351-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        /// <exception cref="NetOfficeCOMException">Unexpected error</exception>
        public _CommandBarButtonEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region _CommandBarButtonEvents

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="cancelDefault"></param>
        public void Click([In, MarshalAs(UnmanagedType.IDispatch)] object ctrl, [In] [Out] ref object cancelDefault)
        {
            if (!Validate("Click"))
            {
                Invoker.ReleaseParamsArray(ctrl, cancelDefault);
                return;
            }

            NetOffice.OfficeApi.CommandBarButton newCtrl = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CommandBarButton>(EventClass, ctrl, typeof(NetOffice.OfficeApi.CommandBarButton));
            object[] paramsArray = new object[2];
            paramsArray[0] = newCtrl;
            paramsArray.SetValue(cancelDefault, 1);
            EventBinding.RaiseCustomEvent("Click", ref paramsArray);

            cancelDefault = ToBoolean(paramsArray[1]);
        }

        #endregion
    }
}
