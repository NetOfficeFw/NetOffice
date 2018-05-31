using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{

    /// <summary>
    /// Interface IChartFillFormat 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public interface IChartFillFormat : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.ChartColorFormat BackColor { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.ChartColorFormat ForeColor { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoGradientColorType GradientColorType { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Single GradientDegree { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoGradientStyle GradientStyle { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 GradientVariant { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoPatternType Pattern { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoPresetGradientType PresetGradientType { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoPresetTexture PresetTexture { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string TextureName { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoTextureType TextureType { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoFillType Type { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoTriState Visible { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="style">NetOffice.OfficeApi.Enums.MsoGradientStyle style</param>
        /// <param name="variant">Int32 variant</param>
        /// <param name="degree">Single degree</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 OneColorGradient(NetOffice.OfficeApi.Enums.MsoGradientStyle style, Int32 variant, Single degree);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="style">NetOffice.OfficeApi.Enums.MsoGradientStyle style</param>
        /// <param name="variant">Int32 variant</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 TwoColorGradient(NetOffice.OfficeApi.Enums.MsoGradientStyle style, Int32 variant);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="presetTexture">NetOffice.OfficeApi.Enums.MsoPresetTexture presetTexture</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 PresetTextured(NetOffice.OfficeApi.Enums.MsoPresetTexture presetTexture);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Solid();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pattern">NetOffice.OfficeApi.Enums.MsoPatternType pattern</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Patterned(NetOffice.OfficeApi.Enums.MsoPatternType pattern);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pictureFile">optional object pictureFile</param>
        /// <param name="pictureFormat">optional object pictureFormat</param>
        /// <param name="pictureStackUnit">optional object pictureStackUnit</param>
        /// <param name="picturePlacement">optional object picturePlacement</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 UserPicture(object pictureFile, object pictureFormat, object pictureStackUnit, object picturePlacement);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 UserPicture();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pictureFile">optional object pictureFile</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 UserPicture(object pictureFile);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pictureFile">optional object pictureFile</param>
        /// <param name="pictureFormat">optional object pictureFormat</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 UserPicture(object pictureFile, object pictureFormat);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pictureFile">optional object pictureFile</param>
        /// <param name="pictureFormat">optional object pictureFormat</param>
        /// <param name="pictureStackUnit">optional object pictureStackUnit</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 UserPicture(object pictureFile, object pictureFormat, object pictureStackUnit);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="textureFile">string textureFile</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 UserTextured(string textureFile);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="style">NetOffice.OfficeApi.Enums.MsoGradientStyle style</param>
        /// <param name="variant">Int32 variant</param>
        /// <param name="presetGradientType">NetOffice.OfficeApi.Enums.MsoPresetGradientType presetGradientType</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 PresetGradient(NetOffice.OfficeApi.Enums.MsoGradientStyle style, Int32 variant, NetOffice.OfficeApi.Enums.MsoPresetGradientType presetGradientType);

        #endregion
    }
}
