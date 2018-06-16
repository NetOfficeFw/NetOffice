using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.MSComctlLibApi.EventContracts.ImageListEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class ImageListEvents_SinkHelper : SinkHelper, NetOffice.MSComctlLibApi.EventContracts.ImageListEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from ImageListEvents
        /// </summary>
        public static readonly string Id = "2C247F22-8591-11D1-B16A-00C0F0283628";

        #endregion

        #region Fields

        //private IEventBinding _eventBinding;
        //private ICOMObject _eventClass;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public ImageListEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion
    }
}
