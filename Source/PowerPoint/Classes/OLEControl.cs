using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OLEControl_GotFocusEventHandler();
	public delegate void OLEControl_LostFocusEventHandler();
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass OLEControl
    /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OCXExtenderEvents))]
	[TypeId("91493446-5A91-11CF-8700-00AA0060263B")]
    public interface OLEControl : OCXExtender, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        event OLEControl_GotFocusEventHandler GotFocusEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        event OLEControl_LostFocusEventHandler LostFocusEvent;

        #endregion
    }
}
