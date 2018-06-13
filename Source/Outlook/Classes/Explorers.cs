using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Explorers_NewExplorerEventHandler(NetOffice.OutlookApi._Explorer explorer);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Explorers 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867227.aspx </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("00063053-0000-0000-C000-000000000046")]
	public interface Explorers : _Explorers, IEventBinding
	{
        #region Events

        /// <summary>
        /// SupportByVersion Outlook 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866400.aspx </remarks>
        [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		event Explorers_NewExplorerEventHandler NewExplorerEvent;

        #endregion
    }
}
