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
    /// Default implementation of <see cref="NetOffice.OfficeApi.EventContracts._CommandBarComboBoxEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class _CommandBarComboBoxEvents_SinkHelper : SinkHelper, NetOffice.OfficeApi.EventContracts._CommandBarComboBoxEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from _CommandBarComboBoxEvents
        /// </summary>
        public static readonly string Id = "000C0354-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        /// <exception cref="NetOfficeCOMException">Unexpected error</exception>
        public _CommandBarComboBoxEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region _CommandBarComboBoxEvents

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ctrl"></param>
        public void Change([In, MarshalAs(UnmanagedType.IDispatch)] object ctrl)
        {
            if (!Validate("Change"))
            {
                Invoker.ReleaseParamsArray(ctrl);
                return;
            }

            NetOffice.OfficeApi.CommandBarComboBox newCtrl = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CommandBarComboBox>(EventClass, ctrl, typeof(NetOffice.OfficeApi.CommandBarComboBox));
            object[] paramsArray = new object[1];
            paramsArray[0] = newCtrl;
            EventBinding.RaiseCustomEvent("Change", ref paramsArray);
        }

        #endregion
    }
}
