using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface LineFormat 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("00021268-0000-0000-C000-000000000046")]
	public interface LineFormat : ICOMObject
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
		NetOffice.PublisherApi.ColorFormat BackColor { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoArrowheadLength BeginArrowheadLength { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoArrowheadStyle BeginArrowheadStyle { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoArrowheadWidth BeginArrowheadWidth { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoLineDashStyle DashStyle { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoArrowheadLength EndArrowheadLength { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoArrowheadStyle EndArrowheadStyle { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoArrowheadWidth EndArrowheadWidth { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ColorFormat ForeColor { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPatternType Pattern { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoLineStyle Style { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Single Transparency { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState Visible { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		object Weight { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState InsetPen { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		Single GradientAngle { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		NetOffice.OfficeApi.Enums.MsoGradientColorType GradientColorType { get; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		NetOffice.OfficeApi.Enums.MsoGradientStyle GradientStyle { get; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		Int32 GradientVariant { get; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		NetOffice.OfficeApi.Enums.MsoPresetGradientType PresetGradientType { get; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		NetOffice.OfficeApi.Enums.MsoLineCapStyle CapStyle { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		NetOffice.OfficeApi.Enums.MsoLineJoinStyle JoinStyle { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		NetOffice.OfficeApi.Enums.MsoLineFillType Type { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// </summary>
		/// <param name="style">NetOffice.OfficeApi.Enums.MsoGradientStyle style</param>
		/// <param name="variant">Int32 variant</param>
		/// <param name="presetGradientType">NetOffice.OfficeApi.Enums.MsoPresetGradientType presetGradientType</param>
		[SupportByVersion("Publisher", 15,16)]
		void PresetGradient(NetOffice.OfficeApi.Enums.MsoGradientStyle style, Int32 variant, NetOffice.OfficeApi.Enums.MsoPresetGradientType presetGradientType);

		#endregion
	}
}
