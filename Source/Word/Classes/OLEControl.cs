using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OLEControl_GotFocusEventHandler();
	public delegate void OLEControl_LostFocusEventHandler();
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass OLEControl 
    /// SupportByVersion Word, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OCXEvents))]
	[TypeId("000209F2-0000-0000-C000-000000000046")]
    public interface OLEControl : _OLEControl, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion Word 9 10 11 12 14 15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        event OLEControl_GotFocusEventHandler GotFocusEvent;

        /// <summary>
        /// SupportByVersion Word 9 10 11 12 14 15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        event OLEControl_LostFocusEventHandler LostFocusEvent;

        #endregion
    }  
}
