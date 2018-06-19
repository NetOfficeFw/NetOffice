using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface SoftEdgeFormat 
	/// SupportByVersion Publisher, 15,16
	/// </summary>
	[SupportByVersion("Publisher", 15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("00021270-0000-0000-C000-000000000046")]
	public interface SoftEdgeFormat : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		NetOffice.OfficeApi.Enums.MsoSoftEdgeType Type { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		Single Radius { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState Visible { get; set; }

		#endregion

	}
}
