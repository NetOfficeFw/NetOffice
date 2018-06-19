using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface ReflectionFormat 
	/// SupportByVersion Publisher, 15,16
	/// </summary>
	[SupportByVersion("Publisher", 15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("00021272-0000-0000-C000-000000000046")]
	public interface ReflectionFormat : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		NetOffice.OfficeApi.Enums.MsoReflectionType Type { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		Single Transparency { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		Single Size { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		Single Offset { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		Single Blur { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState Visible { get; set; }

		#endregion

	}
}
