using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	#region Delegates

	#pragma warning disable
	public delegate void DataRecordsets_DataRecordsetAddedEventHandler(NetOffice.VisioApi.IVDataRecordset dataRecordset);
	public delegate void DataRecordsets_BeforeDataRecordsetDeleteEventHandler(NetOffice.VisioApi.IVDataRecordset dataRecordset);
	public delegate void DataRecordsets_DataRecordsetChangedEventHandler(NetOffice.VisioApi.IVDataRecordsetChangedEvent dataRecordsetChanged);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass DataRecordsets 
	/// SupportByVersion Visio, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769264(v=office.14).aspx </remarks>
	[SupportByVersion("Visio", 12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.VisioApi.EventContracts.EDataRecordsets))]
	[TypeId("000D0A2B-0000-0000-C000-000000000046")]
    public interface DataRecordsets : IVDataRecordsets, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769186(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event DataRecordsets_DataRecordsetAddedEventHandler DataRecordsetAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769026(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event DataRecordsets_BeforeDataRecordsetDeleteEventHandler BeforeDataRecordsetDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767609(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		event DataRecordsets_DataRecordsetChangedEventHandler DataRecordsetChangedEvent;

		#endregion
	}
}
