using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface ThreeDFormat 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("0002126F-0000-0000-C000-000000000046")]
	public interface ThreeDFormat : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		object Depth { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ColorFormat ExtrusionColor { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoExtrusionColorType ExtrusionColorType { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState Perspective { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection PresetExtrusionDirection { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPresetLightingDirection PresetLightingDirection { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPresetLightingSoftness PresetLightingSoftness { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPresetMaterial PresetMaterial { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat PresetThreeDFormat { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Single RotationX { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Single RotationY { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState Visible { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		NetOffice.OfficeApi.Enums.MsoBevelType BevelTopType { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		Single BevelTopInset { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		Single BevelTopDepth { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		NetOffice.OfficeApi.Enums.MsoBevelType BevelBottomType { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		Single BevelBottomInset { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		Single BevelBottomDepth { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		Single ContourWidth { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		NetOffice.PublisherApi.ColorFormat ContourColor { get; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		Single FieldOfView { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void IncrementRotationX(Single increment);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void IncrementRotationY(Single increment);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void ResetRotation();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="presetThreeDFormat">NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat presetThreeDFormat</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void SetThreeDFormat(NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat presetThreeDFormat);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="presetExtrusionDirection">NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection presetExtrusionDirection</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void SetExtrusionDirection(NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection presetExtrusionDirection);

		#endregion
	}
}
