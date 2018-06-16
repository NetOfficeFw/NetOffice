using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface ChInterior 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("56C833A6-3E1C-11D3-831A-00C04F991C70")]
	public interface ChInterior : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		object Color { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		object DefaultColor { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		object BackColor { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartPatternTypeEnum Pattern { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartFillTypeEnum FillType { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartPresetGradientTypeEnum PresetGradientType { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartGradientStyleEnum GradientStyle { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartGradientVariantEnum GradientVariant { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Double GradientDegree { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartPresetTextureEnum PresetTexture { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string TextureName { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartTextureFormatEnum TextureFormat { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Double TextureStackUnit { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.ChartTexturePlacementEnum TexturePlacement { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="patternType">NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType</param>
		/// <param name="color">optional object color</param>
		/// <param name="backColor">optional object backColor</param>
		[SupportByVersion("OWC10", 1)]
		void SetPatterned(NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType, object color, object backColor);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="patternType">NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void SetPatterned(NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="patternType">NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType</param>
		/// <param name="color">optional object color</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void SetPatterned(NetOffice.OWC10Api.Enums.ChartPatternTypeEnum patternType, object color);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle</param>
		/// <param name="gradientVarient">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVarient</param>
		/// <param name="gradientPreset">NetOffice.OWC10Api.Enums.ChartPresetGradientTypeEnum gradientPreset</param>
		[SupportByVersion("OWC10", 1)]
		void SetPresetGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVarient, NetOffice.OWC10Api.Enums.ChartPresetGradientTypeEnum gradientPreset);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="textureFile">object textureFile</param>
		/// <param name="textureFormat">optional NetOffice.OWC10Api.Enums.ChartTextureFormatEnum TextureFormat = 4</param>
		/// <param name="stackUnit">optional Double stackUnit = 0</param>
		/// <param name="texturePlacement">optional NetOffice.OWC10Api.Enums.ChartTexturePlacementEnum TexturePlacement = 7</param>
		[SupportByVersion("OWC10", 1)]
		void SetTextured(object textureFile, object textureFormat, object stackUnit, object texturePlacement);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="textureFile">object textureFile</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void SetTextured(object textureFile);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="textureFile">object textureFile</param>
		/// <param name="textureFormat">optional NetOffice.OWC10Api.Enums.ChartTextureFormatEnum TextureFormat = 4</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void SetTextured(object textureFile, object textureFormat);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="textureFile">object textureFile</param>
		/// <param name="textureFormat">optional NetOffice.OWC10Api.Enums.ChartTextureFormatEnum TextureFormat = 4</param>
		/// <param name="stackUnit">optional Double stackUnit = 0</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void SetTextured(object textureFile, object textureFormat, object stackUnit);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant</param>
		/// <param name="gradientDegree">Double gradientDegree</param>
		/// <param name="color">optional object color</param>
		[SupportByVersion("OWC10", 1)]
		void SetOneColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant, Double gradientDegree, object color);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant</param>
		/// <param name="gradientDegree">Double gradientDegree</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void SetOneColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant, Double gradientDegree);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="color">optional object color</param>
		[SupportByVersion("OWC10", 1)]
		void SetSolid(object color);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void SetSolid();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant</param>
		/// <param name="color">optional object color</param>
		/// <param name="backColor">optional object backColor</param>
		[SupportByVersion("OWC10", 1)]
		void SetTwoColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant, object color, object backColor);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void SetTwoColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="gradientStyle">NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle</param>
		/// <param name="gradientVariant">NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant</param>
		/// <param name="color">optional object color</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void SetTwoColorGradient(NetOffice.OWC10Api.Enums.ChartGradientStyleEnum gradientStyle, NetOffice.OWC10Api.Enums.ChartGradientVariantEnum gradientVariant, object color);

		#endregion
	}
}
