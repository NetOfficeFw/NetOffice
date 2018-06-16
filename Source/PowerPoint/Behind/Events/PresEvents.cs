using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.PowerPointApi.EventContracts.PresEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class PresEvents_SinkHelper : SinkHelper, NetOffice.PowerPointApi.EventContracts.PresEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from PresEvents
        /// </summary>
        public static readonly string Id = "91493463-5A91-11CF-8700-00AA0060263B";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public PresEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion
    }
}
