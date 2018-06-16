using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OWC10Api.EventContracts._NavigationEvent"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class _NavigationEvent_SinkHelper : SinkHelper, NetOffice.OWC10Api.EventContracts._NavigationEvent
    {
        #region Static

        /// <summary>
        /// Interface Id from _NavigationEvent
        /// </summary>
        public static readonly string Id = "4BD09D02-45CC-11D1-B1D1-006097C97F9B";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public _NavigationEvent_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region _NavigationEvent

        /// <summary>
        /// 
        /// </summary>
        /// <param name="navButton"></param>
        public void ButtonClick([In] object navButton)
        {
            if (!Validate("ButtonClick"))
            {
                Invoker.ReleaseParamsArray(navButton);
                return;
            }

            NetOffice.OWC10Api.Enums.NavButtonEnum newNavButton = (NetOffice.OWC10Api.Enums.NavButtonEnum)navButton;
            object[] paramsArray = new object[1];
            paramsArray[0] = newNavButton;
            EventBinding.RaiseCustomEvent("ButtonClick", ref paramsArray);
        }

        #endregion
    }
}
