using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface ChartFillFormat 
	/// SupportByVersion PowerPoint, 14,15,16
	/// </summary>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("92D41A5B-F07E-4CA4-AF6F-BEF486AA4E6F")]
	public interface ChartFillFormat : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.ChartColorFormat BackColor { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.ChartColorFormat ForeColor { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Single GradientDegree { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		string TextureName { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoGradientColorType GradientColorType { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoGradientStyle GradientStyle { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		Int32 GradientVariant { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPatternType Pattern { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPresetGradientType PresetGradientType { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPresetTexture PresetTexture { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTextureType TextureType { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoFillType Type { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState Visible { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void Solid();

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="textureFile">string textureFile</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void UserTextured(string textureFile);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="style">NetOffice.OfficeApi.Enums.MsoGradientStyle style</param>
		/// <param name="variant">Int32 variant</param>
		/// <param name="degree">Single degree</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void OneColorGradient(NetOffice.OfficeApi.Enums.MsoGradientStyle style, Int32 variant, Single degree);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="pattern">NetOffice.OfficeApi.Enums.MsoPatternType pattern</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void Patterned(NetOffice.OfficeApi.Enums.MsoPatternType pattern);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="style">NetOffice.OfficeApi.Enums.MsoGradientStyle style</param>
		/// <param name="variant">Int32 variant</param>
		/// <param name="presetGradientType">NetOffice.OfficeApi.Enums.MsoPresetGradientType presetGradientType</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void PresetGradient(NetOffice.OfficeApi.Enums.MsoGradientStyle style, Int32 variant, NetOffice.OfficeApi.Enums.MsoPresetGradientType presetGradientType);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="presetTexture">NetOffice.OfficeApi.Enums.MsoPresetTexture presetTexture</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void PresetTextured(NetOffice.OfficeApi.Enums.MsoPresetTexture presetTexture);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="style">NetOffice.OfficeApi.Enums.MsoGradientStyle style</param>
		/// <param name="variant">Int32 variant</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void TwoColorGradient(NetOffice.OfficeApi.Enums.MsoGradientStyle style, Int32 variant);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="pictureFile">optional object pictureFile</param>
		/// <param name="pictureFormat">optional object pictureFormat</param>
		/// <param name="pictureStackUnit">optional object pictureStackUnit</param>
		/// <param name="picturePlacement">optional object picturePlacement</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void UserPicture(object pictureFile, object pictureFormat, object pictureStackUnit, object picturePlacement);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		void UserPicture();

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="pictureFile">optional object pictureFile</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		void UserPicture(object pictureFile);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="pictureFile">optional object pictureFile</param>
		/// <param name="pictureFormat">optional object pictureFormat</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		void UserPicture(object pictureFile, object pictureFormat);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="pictureFile">optional object pictureFile</param>
		/// <param name="pictureFormat">optional object pictureFormat</param>
		/// <param name="pictureStackUnit">optional object pictureStackUnit</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		void UserPicture(object pictureFile, object pictureFormat, object pictureStackUnit);

		#endregion
	}
}
