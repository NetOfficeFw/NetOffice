using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OLEObject_GotFocusEventHandler();
	public delegate void OLEObject_LostFocusEventHandler();
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass OLEObject
    /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838421.aspx </remarks>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.OLEObjectEvents))]
	[TypeId("00020818-0000-0000-C000-000000000046")]
    public interface OLEObject : _OLEObject, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195806.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event OLEObject_GotFocusEventHandler GotFocusEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836813.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        event OLEObject_LostFocusEventHandler LostFocusEvent;

        #endregion
    }
}
