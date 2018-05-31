using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// Interface IListDataFormat 
    /// SupportByVersion Excel, 11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public interface IListDataFormat : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlListDataType _Default { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        object Choices { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        Int32 DecimalPlaces { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        object DefaultValue { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        bool IsPercent { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        Int32 lcid { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        Int32 MaxCharacters { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        object MaxNumber { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        object MinNumber { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        bool Required { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlListDataType Type { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        bool ReadOnly { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        bool AllowFillIn { get; }

        #endregion
    }
}
