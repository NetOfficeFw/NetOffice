using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface ThreeDFormat 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836783.aspx </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
    [Duplicate("NetOffice.OfficeApi.ThreeDFormat")]
	public interface ThreeDFormat : NetOffice.OfficeApi._IMsoDispObj
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196683.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194958.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Single Depth { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839765.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.ColorFormat ExtrusionColor { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839061.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoExtrusionColorType ExtrusionColorType { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837061.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState Perspective { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821812.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection PresetExtrusionDirection { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821262.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPresetLightingDirection PresetLightingDirection { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840321.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPresetLightingSoftness PresetLightingSoftness { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841150.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPresetMaterial PresetMaterial { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822170.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat PresetThreeDFormat { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840434.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Single RotationX { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822874.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Single RotationY { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821252.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState Visible { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822331.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoLightRigType PresetLighting { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834697.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		Single Z { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194951.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoBevelType BevelTopType { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838646.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		Single BevelTopInset { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197277.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		Single BevelTopDepth { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821639.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoBevelType BevelBottomType { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196509.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		Single BevelBottomInset { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835271.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		Single BevelBottomDepth { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193860.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPresetCamera PresetCamera { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198152.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		Single RotationZ { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194602.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		Single ContourWidth { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836496.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.ColorFormat ContourColor { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822572.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		Single FieldOfView { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838170.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState ProjectText { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193315.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		Single LightAngle { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821206.aspx </remarks>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void IncrementRotationX(Single increment);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820944.aspx </remarks>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void IncrementRotationY(Single increment);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820893.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void ResetRotation();

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821911.aspx </remarks>
		/// <param name="presetThreeDFormat">NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat presetThreeDFormat</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void SetThreeDFormat(NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat presetThreeDFormat);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196541.aspx </remarks>
		/// <param name="presetExtrusionDirection">NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection presetExtrusionDirection</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		void SetExtrusionDirection(NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection presetExtrusionDirection);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820808.aspx </remarks>
		/// <param name="presetCamera">NetOffice.OfficeApi.Enums.MsoPresetCamera presetCamera</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		void SetPresetCamera(NetOffice.OfficeApi.Enums.MsoPresetCamera presetCamera);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196293.aspx </remarks>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		void IncrementRotationZ(Single increment);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196024.aspx </remarks>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		void IncrementRotationHorizontal(Single increment);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193527.aspx </remarks>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		void IncrementRotationVertical(Single increment);

		#endregion
	}
}
