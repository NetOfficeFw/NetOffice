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
    public class _CommandBarsEvents_SinkHelper : SinkHelper, NetOffice.OfficeApi.EventContracts._CommandBarsEvents
    {
        #region Static

        public static readonly string Id = "000C0352-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        public _CommandBarsEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region _CommandBarsEvents

        public void OnUpdate()
        {
            if (!Validate("OnUpdate"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("OnUpdate", ref paramsArray);
        }

        #endregion
    }
}
