using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.Behind.EventContracts
{
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class _CommandBarButtonEvents_SinkHelper : SinkHelper, NetOffice.OfficeApi.EventContracts._CommandBarButtonEvents
    {
        #region Static

        public static readonly string Id = "000C0351-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        public _CommandBarButtonEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region _CommandBarButtonEvents

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
