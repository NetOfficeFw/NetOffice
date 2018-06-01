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
    /// Default implementation of <see cref="NetOffice.OfficeApi.EventContracts.IMsoEnvelopeVBEvents"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class IMsoEnvelopeVBEvents_SinkHelper : SinkHelper, NetOffice.OfficeApi.EventContracts.IMsoEnvelopeVBEvents
    {
        #region Static

        /// <summary>
        /// Interface Id from IMsoEnvelopeVBEvents
        /// </summary>
        public static readonly string Id = "000672AD-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        /// <exception cref="NetOfficeCOMException">Unexpected error</exception>
        public IMsoEnvelopeVBEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region IMsoEnvelopeVBEvents

        /// <summary>
        /// 
        /// </summary>
        public void EnvelopeShow()
        {
            if (!Validate("EnvelopeShow"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("EnvelopeShow", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        public void EnvelopeHide()
        {
            if (!Validate("EnvelopeHide"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("EnvelopeHide", ref paramsArray);
        }

        #endregion
    }
}
