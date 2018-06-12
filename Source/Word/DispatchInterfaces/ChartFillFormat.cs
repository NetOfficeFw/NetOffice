using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface ChartFillFormat 
	/// SupportByVersion Word, 14,15,16
	/// </summary>
	[SupportByVersion("Word", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface ChartFillFormat : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.ChartColorFormat BackColor { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.ChartColorFormat ForeColor { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoGradientColorType GradientColorType { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		Single GradientDegree { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoGradientStyle GradientStyle { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		Int32 GradientVariant { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPatternType Pattern { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPresetGradientType PresetGradientType { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoPresetTexture PresetTexture { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		string TextureName { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTextureType TextureType { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoFillType Type { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState Visible { get; set; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		object Application { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		object Parent { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="style">NetOffice.OfficeApi.Enums.MsoGradientStyle style</param>
		/// <param name="variant">Int32 variant</param>
		/// <param name="degree">Single degree</param>
		[SupportByVersion("Word", 14,15,16)]
		void OneColorGradient(NetOffice.OfficeApi.Enums.MsoGradientStyle style, Int32 variant, Single degree);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="style">NetOffice.OfficeApi.Enums.MsoGradientStyle style</param>
		/// <param name="variant">Int32 variant</param>
		[SupportByVersion("Word", 14,15,16)]
		void TwoColorGradient(NetOffice.OfficeApi.Enums.MsoGradientStyle style, Int32 variant);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="presetTexture">NetOffice.OfficeApi.Enums.MsoPresetTexture presetTexture</param>
		[SupportByVersion("Word", 14,15,16)]
		void PresetTextured(NetOffice.OfficeApi.Enums.MsoPresetTexture presetTexture);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		void Solid();

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="pattern">NetOffice.OfficeApi.Enums.MsoPatternType pattern</param>
		[SupportByVersion("Word", 14,15,16)]
		void Patterned(NetOffice.OfficeApi.Enums.MsoPatternType pattern);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="pictureFile">optional object pictureFile</param>
		/// <param name="pictureFormat">optional object pictureFormat</param>
		/// <param name="pictureStackUnit">optional object pictureStackUnit</param>
		/// <param name="picturePlacement">optional object picturePlacement</param>
		[SupportByVersion("Word", 14,15,16)]
		void UserPicture(object pictureFile, object pictureFormat, object pictureStackUnit, object picturePlacement);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void UserPicture();

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="pictureFile">optional object pictureFile</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void UserPicture(object pictureFile);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="pictureFile">optional object pictureFile</param>
		/// <param name="pictureFormat">optional object pictureFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void UserPicture(object pictureFile, object pictureFormat);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="pictureFile">optional object pictureFile</param>
		/// <param name="pictureFormat">optional object pictureFormat</param>
		/// <param name="pictureStackUnit">optional object pictureStackUnit</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		void UserPicture(object pictureFile, object pictureFormat, object pictureStackUnit);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="textureFile">string textureFile</param>
		[SupportByVersion("Word", 14,15,16)]
		void UserTextured(string textureFile);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="style">NetOffice.OfficeApi.Enums.MsoGradientStyle style</param>
		/// <param name="variant">Int32 variant</param>
		/// <param name="presetGradientType">NetOffice.OfficeApi.Enums.MsoPresetGradientType presetGradientType</param>
		[SupportByVersion("Word", 14,15,16)]
		void PresetGradient(NetOffice.OfficeApi.Enums.MsoGradientStyle style, Int32 variant, NetOffice.OfficeApi.Enums.MsoPresetGradientType presetGradientType);

		#endregion
	}
}
