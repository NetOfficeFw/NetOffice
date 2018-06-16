using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.OWC10Api.EventContracts.IRangeEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class IRangeEvents_SinkHelper : SinkHelper, NetOffice.OWC10Api.EventContracts.IRangeEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from IRangeEvents
        /// </summary>
        public static readonly string Id = "B8891063-2B00-48EC-957F-6DEBEADE9D8B";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public IRangeEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region IRangeEvents

        /// <summary>
        /// 
        /// </summary>
        public void Change()
        {
            if (!Validate("Change"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Change", ref paramsArray);
        }

        #endregion
    }
}
