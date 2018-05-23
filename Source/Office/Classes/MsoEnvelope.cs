using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	#region Delegates

	#pragma warning disable
	public delegate void MsoEnvelope_EnvelopeShowEventHandler();
	public delegate void MsoEnvelope_EnvelopeHideEventHandler();
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass MsoEnvelope
    /// SupportByVersion Office, 10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862112.aspx </remarks>
    [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventInterface(typeof(NetOffice.OfficeApi.EventInterfaces.IMsoEnvelopeVBEvents))]
    public interface MsoEnvelope : IMsoEnvelopeVB, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion Office 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861098.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        event MsoEnvelope_EnvelopeShowEventHandler EnvelopeShowEvent;

        /// <summary>
        /// SupportByVersion Office 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860254.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        event MsoEnvelope_EnvelopeHideEventHandler EnvelopeHideEvent;

        #endregion
    }
}
