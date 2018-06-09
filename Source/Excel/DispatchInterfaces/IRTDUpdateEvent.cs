using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// DispatchInterface IRTDUpdateEvent 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("A43788C1-D91B-11D3-8F39-00C04F3651B8")]
	public interface IRTDUpdateEvent : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks>Setting the HeartbeatInterval property to -1 will result in the Heartbeat method not being called. Note The heartbeat interval cannot be set below 15,000 milliseconds due to the standard 15-second time out</remarks>
		[SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
		Int32 HeartbeatInterval { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
		void UpdateNotify();

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
		void Disconnect();

		#endregion
	}
}
