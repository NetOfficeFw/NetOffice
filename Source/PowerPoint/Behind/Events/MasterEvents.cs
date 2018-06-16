using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.PowerPointApi.EventContracts.MasterEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class MasterEvents_SinkHelper : SinkHelper, NetOffice.PowerPointApi.EventContracts.MasterEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from MasterEvents
        /// </summary>
        public static readonly string Id = "914934D2-5A91-11CF-8700-00AA0060263B";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public MasterEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion
    }
}
