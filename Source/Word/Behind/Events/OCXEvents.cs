using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.WordApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.WordApi.EventContracts.OCXEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class OCXEvents_SinkHelper : SinkHelper, NetOffice.WordApi.EventContracts.OCXEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from OCXEvents
        /// </summary>
        public static readonly string Id = "000209F3-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public OCXEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region OCXEvents 

        /// <summary>
        /// 
        /// </summary>
        public void GotFocus()
        {
            if (!Validate("GotFocus"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("GotFocus", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        public void LostFocus()
        {
            if (!Validate("LostFocus"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("LostFocus", ref paramsArray);
        }

        #endregion
    }
}
