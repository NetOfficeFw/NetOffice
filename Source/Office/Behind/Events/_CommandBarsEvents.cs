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
    /// Default implementation of <see cref="NetOffice.OfficeApi.EventContracts._CommandBarsEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class _CommandBarsEvents_SinkHelper : SinkHelper, NetOffice.OfficeApi.EventContracts._CommandBarsEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from _CommandBarsEvents
        /// </summary>
        public static readonly string Id = "000C0352-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        /// <exception cref="NetOfficeCOMException">Unexpected error</exception>
        public _CommandBarsEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region _CommandBarsEvents

        /// <summary>
        /// 
        /// </summary>
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
