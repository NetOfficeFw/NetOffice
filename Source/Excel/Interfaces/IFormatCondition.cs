using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// Interface IFormatCondition 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public interface IFormatCondition : ICOMObject
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
        Int32 Type { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Operator { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Formula1 { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Formula2 { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Interior Interior { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Borders Borders { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Font Font { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Text { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlContainsOperator TextOperator { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlTimePeriods DateOperator { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        object NumberFormat { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Int32 Priority { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool StopIfTrue { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range AppliesTo { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool PTCondition { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlPivotConditionScope ScopeType { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="formula1">optional object formula1</param>
        /// <param name="formula2">optional object formula2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="formula1">optional object formula1</param>
        /// <param name="formula2">optional object formula2</param>
        /// <param name="_string">optional object string</param>
        /// <param name="operator2">optional object operator2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Int32 Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2, object _string, object operator2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="formula1">optional object formula1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="formula1">optional object formula1</param>
        /// <param name="formula2">optional object formula2</param>
        /// <param name="_string">optional object string</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Int32 Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2, object _string);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Delete();

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="formula1">optional object formula1</param>
        /// <param name="formula2">optional object formula2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Int32 _Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Int32 _Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Int32 _Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="formula1">optional object formula1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Int32 _Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="range">NetOffice.ExcelApi.Range range</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Int32 ModifyAppliesToRange(NetOffice.ExcelApi.Range range);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Int32 SetFirstPriority();

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Int32 SetLastPriority();

        #endregion
    }
}
