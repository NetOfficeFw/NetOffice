using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// Interface IConditionValue 
    /// SupportByVersion Excel, 12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public interface IConditionValue : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlConditionValueTypes Type { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        object Value { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="newtype">NetOffice.ExcelApi.Enums.XlConditionValueTypes newtype</param>
        /// <param name="newvalue">optional object newvalue</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Int32 Modify(NetOffice.ExcelApi.Enums.XlConditionValueTypes newtype, object newvalue);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="newtype">NetOffice.ExcelApi.Enums.XlConditionValueTypes newtype</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Int32 Modify(NetOffice.ExcelApi.Enums.XlConditionValueTypes newtype);

        #endregion
    }
}
