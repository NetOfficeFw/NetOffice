using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Inspectors_NewInspectorEventHandler(NetOffice.OutlookApi._Inspector inspector);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Inspectors 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868697.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.InspectorsEvents))]
	[TypeId("00063054-0000-0000-C000-000000000046")]
    public interface Inspectors : _Inspectors, IEventBinding
    { 
        #region Events
        
        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867841.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Inspectors_NewInspectorEventHandler NewInspectorEvent;

        #endregion
    }
}
