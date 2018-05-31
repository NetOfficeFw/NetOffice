using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi.Behind.EventContracts
{
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class RefreshEvents_SinkHelper : SinkHelper, NetOffice.ExcelApi.EventContracts.RefreshEvents
    {
        #region Static

        public static readonly string Id = "0002441B-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        public RefreshEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region RefreshEvents

        public void BeforeRefresh([In] [Out] ref object cancel)
        {
            if (!Validate("BeforeRefresh"))
            {
                Invoker.ReleaseParamsArray(cancel);
                return;
            }

            object[] paramsArray = new object[1];
            paramsArray.SetValue(cancel, 0);
            EventBinding.RaiseCustomEvent("BeforeRefresh", ref paramsArray);

            cancel = ToBoolean(paramsArray[0]);
        }

        public void AfterRefresh([In] object success)
        {
            if (!Validate("AfterRefresh"))
            {
                Invoker.ReleaseParamsArray(success);
                return;
            }

            bool newSuccess = ToBoolean(success);
            object[] paramsArray = new object[1];
            paramsArray[0] = newSuccess;
            EventBinding.RaiseCustomEvent("AfterRefresh", ref paramsArray);
        }

        #endregion
    }
}
