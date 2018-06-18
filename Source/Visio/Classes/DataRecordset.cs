using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	#region Delegates

	#pragma warning disable
	public delegate void DataRecordset_DataRecordsetChangedEventHandler(NetOffice.VisioApi.IVDataRecordsetChangedEvent dataRecordsetChanged);
	public delegate void DataRecordset_BeforeDataRecordsetDeleteEventHandler(NetOffice.VisioApi.IVDataRecordset dataRecordset);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass DataRecordset 
	/// SupportByVersion Visio, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769258(v=office.14).aspx </remarks>
	[SupportByVersion("Visio", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.VisioApi.EventContracts.EDataRecordset))]
	[TypeId("000D0A2C-0000-0000-C000-000000000046")]
    public interface DataRecordset : IVDataRecordset, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766153(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event DataRecordset_DataRecordsetChangedEventHandler DataRecordsetChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766834(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event DataRecordset_BeforeDataRecordsetDeleteEventHandler BeforeDataRecordsetDeleteEvent;

		#endregion
	}
}
