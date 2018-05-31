using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// Interface IWorksheetFunction 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public partial interface IWorksheetFunction : ICOMObject
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

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool IsNA(object arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool IsError(object arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Dollar(Double arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Dollar(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Fixed(Double arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Fixed(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Fixed(Double arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Pi();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Ln(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Log10(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Round(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Lookup(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Lookup(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Index(object arg1, Double arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Index(object arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Index(object arg1, Double arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Rept(string arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double DCount(NetOffice.ExcelApi.Range arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double DSum(NetOffice.ExcelApi.Range arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double DAverage(NetOffice.ExcelApi.Range arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double DMin(NetOffice.ExcelApi.Range arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double DMax(NetOffice.ExcelApi.Range arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double DStDev(NetOffice.ExcelApi.Range arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double DVar(NetOffice.ExcelApi.Range arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">string arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Text(object arg1, string arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object LinEst(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object LinEst(object arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object LinEst(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object LinEst(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Trend(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Trend(object arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Trend(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Trend(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object LogEst(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object LogEst(object arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object LogEst(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object LogEst(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Growth(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Growth(object arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Growth(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Growth(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Pv(Double arg1, Double arg2, Double arg3, object arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Pv(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Pv(Double arg1, Double arg2, Double arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Fv(Double arg1, Double arg2, Double arg3, object arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Fv(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Fv(Double arg1, Double arg2, Double arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double NPer(Double arg1, Double arg2, Double arg3, object arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double NPer(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double NPer(Double arg1, Double arg2, Double arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Pmt(Double arg1, Double arg2, Double arg3, object arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Pmt(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Pmt(Double arg1, Double arg2, Double arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Rate(Double arg1, Double arg2, Double arg3, object arg4, object arg5, object arg6);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Rate(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Rate(Double arg1, Double arg2, Double arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Rate(Double arg1, Double arg2, Double arg3, object arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double MIrr(object arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Irr(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Irr(object arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Match(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Match(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Weekday(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Weekday(object arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Search(string arg1, string arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Search(string arg1, string arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Transpose(object arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Atan2(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Asin(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Acos(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object HLookup(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object HLookup(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object VLookup(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object VLookup(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Log(Double arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Log(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Proper(string arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Trim(string arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">string arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Replace(string arg1, Double arg2, Double arg3, string arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        /// <param name="arg3">string arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Substitute(string arg1, string arg2, string arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        /// <param name="arg3">string arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Substitute(string arg1, string arg2, string arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Find(string arg1, string arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Find(string arg1, string arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool IsErr(object arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool IsText(object arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool IsNumber(object arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Sln(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Syd(Double arg1, Double arg2, Double arg3, Double arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Ddb(Double arg1, Double arg2, Double arg3, Double arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Ddb(Double arg1, Double arg2, Double arg3, Double arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Clean(string arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double MDeterm(object arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object MInverse(object arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object MMult(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Ipmt(Double arg1, Double arg2, Double arg3, Double arg4, object arg5, object arg6);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Ipmt(Double arg1, Double arg2, Double arg3, Double arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Ipmt(Double arg1, Double arg2, Double arg3, Double arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Ppmt(Double arg1, Double arg2, Double arg3, Double arg4, object arg5, object arg6);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Ppmt(Double arg1, Double arg2, Double arg3, Double arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Ppmt(Double arg1, Double arg2, Double arg3, Double arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Fact(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double DProduct(NetOffice.ExcelApi.Range arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool IsNonText(object arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double DStDevP(NetOffice.ExcelApi.Range arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double DVarP(NetOffice.ExcelApi.Range arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool IsLogical(object arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double DCountA(NetOffice.ExcelApi.Range arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string USDollar(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double FindB(string arg1, string arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double FindB(string arg1, string arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double SearchB(string arg1, string arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double SearchB(string arg1, string arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">string arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string ReplaceB(string arg1, Double arg2, Double arg3, string arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double RoundUp(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double RoundDown(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Rank(Double arg1, NetOffice.ExcelApi.Range arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Rank(Double arg1, NetOffice.ExcelApi.Range arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Days360(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Days360(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">Double arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Vdb(Double arg1, Double arg2, Double arg3, Double arg4, Double arg5, object arg6, object arg7);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">Double arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Vdb(Double arg1, Double arg2, Double arg3, Double arg4, Double arg5);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">Double arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Vdb(Double arg1, Double arg2, Double arg3, Double arg4, Double arg5, object arg6);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Sinh(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Cosh(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Tanh(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Asinh(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Acosh(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Atanh(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object DGet(NetOffice.ExcelApi.Range arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Db(Double arg1, Double arg2, Double arg3, Double arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Db(Double arg1, Double arg2, Double arg3, Double arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Frequency(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double BetaDist(Double arg1, Double arg2, Double arg3, object arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double BetaDist(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double BetaDist(Double arg1, Double arg2, Double arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double GammaLn(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double BetaInv(Double arg1, Double arg2, Double arg3, object arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double BetaInv(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double BetaInv(Double arg1, Double arg2, Double arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double BinomDist(Double arg1, Double arg2, Double arg3, bool arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double ChiDist(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double ChiInv(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Combin(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Confidence(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double CritBinom(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Even(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">bool arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double ExponDist(Double arg1, Double arg2, bool arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double FDist(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double FInv(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Fisher(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double FisherInv(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Floor(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double GammaDist(Double arg1, Double arg2, Double arg3, bool arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double GammaInv(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Ceiling(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double HypGeomDist(Double arg1, Double arg2, Double arg3, Double arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double LogNormDist(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double LogInv(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double NegBinomDist(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double NormDist(Double arg1, Double arg2, Double arg3, bool arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double NormSDist(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double NormInv(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double NormSInv(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Standardize(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Odd(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Permut(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">bool arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Poisson(Double arg1, Double arg2, bool arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double TDist(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Weibull(Double arg1, Double arg2, Double arg3, bool arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double SumXMY2(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double SumX2MY2(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double SumX2PY2(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double ChiTest(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Correl(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Covar(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Forecast(Double arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double FTest(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Intercept(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Pearson(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double RSq(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double StEyx(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Slope(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double TTest(object arg1, object arg2, Double arg3, Double arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Prob(object arg1, object arg2, Double arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Prob(object arg1, object arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double ZTest(object arg1, Double arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double ZTest(object arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Large(object arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Small(object arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Quartile(object arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Percentile(object arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double PercentRank(object arg1, Double arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double PercentRank(object arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double TrimMean(object arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double TInv(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Power(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Radians(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Degrees(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double SumIf(NetOffice.ExcelApi.Range arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double SumIf(NetOffice.ExcelApi.Range arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double CountIf(NetOffice.ExcelApi.Range arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double CountBlank(NetOffice.ExcelApi.Range arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double Ispmt(Double arg1, Double arg2, Double arg3, Double arg4);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Roman(Double arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Roman(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Asc(string arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Dbcs(string arg1);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Phonetic(NetOffice.ExcelApi.Range arg1);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        string BahtText(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        string ThaiDayOfWeek(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        string ThaiDigit(string arg1);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        string ThaiMonthOfYear(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        string ThaiNumSound(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        string ThaiNumString(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        Double ThaiStringLength(string arg1);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool IsThaiDigit(string arg1);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        Double RoundBahtDown(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        Double RoundBahtUp(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        Double ThaiYear(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Hex2Bin(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Hex2Bin(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Hex2Dec(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Hex2Oct(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Hex2Oct(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Dec2Bin(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Dec2Bin(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Dec2Hex(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Dec2Hex(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Dec2Oct(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Dec2Oct(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Oct2Bin(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Oct2Bin(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Oct2Hex(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Oct2Hex(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Oct2Dec(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Bin2Dec(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Bin2Oct(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Bin2Oct(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Bin2Hex(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Bin2Hex(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string ImSub(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string ImDiv(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string ImPower(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string ImAbs(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string ImSqrt(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string ImLn(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string ImLog2(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string ImLog10(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string ImSin(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string ImCos(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string ImExp(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string ImArgument(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string ImConjugate(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Imaginary(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double ImReal(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Complex(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Complex(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double SeriesSum(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double FactDouble(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double SqrtPi(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Quotient(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Delta(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Delta(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double GeStep(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double GeStep(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool IsEven(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool IsOdd(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double MRound(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Erf(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Erf(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double ErfC(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double BesselJ(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double BesselK(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double BesselY(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double BesselI(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Xirr(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Xirr(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Xnpv(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double PriceMat(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double PriceMat(object arg1, object arg2, object arg3, object arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double YieldMat(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double YieldMat(object arg1, object arg2, object arg3, object arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double IntRate(object arg1, object arg2, object arg3, object arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double IntRate(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Received(object arg1, object arg2, object arg3, object arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Received(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Disc(object arg1, object arg2, object arg3, object arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Disc(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double PriceDisc(object arg1, object arg2, object arg3, object arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double PriceDisc(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double YieldDisc(object arg1, object arg2, object arg3, object arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double YieldDisc(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double TBillEq(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double TBillEq(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double TBillPrice(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double TBillPrice(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double TBillYield(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double TBillYield(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Price(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Price(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double DollarDe(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double DollarFr(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Nominal(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Effect(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double CumPrinc(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double CumIPmt(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double EDate(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double EoMonth(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double YearFrac(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double YearFrac(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double CoupDayBs(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double CoupDayBs(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double CoupDays(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double CoupDays(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double CoupDaysNc(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double CoupDaysNc(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double CoupNcd(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double CoupNcd(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double CoupNum(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double CoupNum(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double CoupPcd(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double CoupPcd(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Duration(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Duration(object arg1, object arg2, object arg3, object arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double MDuration(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double MDuration(object arg1, object arg2, object arg3, object arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double OddLPrice(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">object arg7</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double OddLPrice(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double OddLYield(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">object arg7</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double OddLYield(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">object arg7</param>
        /// <param name="arg8">object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double OddFPrice(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">object arg7</param>
        /// <param name="arg8">object arg8</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double OddFPrice(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">object arg7</param>
        /// <param name="arg8">object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double OddFYield(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">object arg7</param>
        /// <param name="arg8">object arg8</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double OddFYield(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double RandBetween(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double WeekNum(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double WeekNum(object arg1);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double AmorDegrc(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double AmorDegrc(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double AmorLinc(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double AmorLinc(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double Convert(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double AccrInt(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double AccrInt(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double AccrIntM(object arg1, object arg2, object arg3, object arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double AccrIntM(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double WorkDay(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double WorkDay(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double NetworkDays(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double NetworkDays(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        Double FVSchedule(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        object IfError(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Confidence_Norm(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Confidence_T(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double ChiSq_Test(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double F_Test(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Covariance_P(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Covariance_S(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">bool arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Expon_Dist(Double arg1, Double arg2, bool arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Gamma_Dist(Double arg1, Double arg2, Double arg3, bool arg4);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Gamma_Inv(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Norm_Dist(Double arg1, Double arg2, Double arg3, bool arg4);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Norm_Inv(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Percentile_Exc(object arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Percentile_Inc(object arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double PercentRank_Exc(object arg1, Double arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        Double PercentRank_Exc(object arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double PercentRank_Inc(object arg1, Double arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        Double PercentRank_Inc(object arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">bool arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Poisson_Dist(Double arg1, Double arg2, bool arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Quartile_Exc(object arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Quartile_Inc(object arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Rank_Avg(Double arg1, NetOffice.ExcelApi.Range arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Rank_Avg(Double arg1, NetOffice.ExcelApi.Range arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Rank_Eq(Double arg1, NetOffice.ExcelApi.Range arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Rank_Eq(Double arg1, NetOffice.ExcelApi.Range arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">bool arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double T_Dist(Double arg1, Double arg2, bool arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double T_Dist_2T(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double T_Dist_RT(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double T_Inv(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double T_Inv_2T(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Weibull_Dist(Double arg1, Double arg2, Double arg3, bool arg4);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double NetworkDays_Intl(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        Double NetworkDays_Intl(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        Double NetworkDays_Intl(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double WorkDay_Intl(object arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        Double WorkDay_Intl(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        Double WorkDay_Intl(object arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double ISO_Ceiling(Double arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        Double ISO_Ceiling(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Dummy21(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Beta_Dist(Double arg1, Double arg2, Double arg3, bool arg4, object arg5, object arg6);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Beta_Dist(Double arg1, Double arg2, Double arg3, bool arg4);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Beta_Dist(Double arg1, Double arg2, Double arg3, bool arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Beta_Inv(Double arg1, Double arg2, Double arg3, object arg4, object arg5);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Beta_Inv(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Beta_Inv(Double arg1, Double arg2, Double arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">bool arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double ChiSq_Dist(Double arg1, Double arg2, bool arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double ChiSq_Dist_RT(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double ChiSq_Inv(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double ChiSq_Inv_RT(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double F_Dist(Double arg1, Double arg2, Double arg3, bool arg4);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double F_Dist_RT(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double F_Inv(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double F_Inv_RT(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">bool arg5</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double HypGeom_Dist(Double arg1, Double arg2, Double arg3, Double arg4, bool arg5);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double LogNorm_Dist(Double arg1, Double arg2, Double arg3, bool arg4);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double LogNorm_Inv(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double NegBinom_Dist(Double arg1, Double arg2, Double arg3, bool arg4);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">bool arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Norm_S_Dist(Double arg1, bool arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Norm_S_Inv(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double T_Test(object arg1, object arg2, Double arg3, Double arg4);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Z_Test(object arg1, Double arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Z_Test(object arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Binom_Dist(Double arg1, Double arg2, Double arg3, bool arg4);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Binom_Inv(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Erf_Precise(object arg1);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double ErfC_Precise(object arg1);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double GammaLn_Precise(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Ceiling_Precise(Double arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Ceiling_Precise(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Floor_Precise(Double arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        Double Floor_Precise(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Acot(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Acoth(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Cot(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Coth(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Csc(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Csch(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Sec(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Sech(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        string ImCot(object arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        string ImTan(object arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        string ImCsc(object arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        string ImCsch(object arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        string ImSec(object arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        string ImSech(object arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Bitand(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Bitor(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Bitxor(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Bitlshift(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Bitrshift(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Combina(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Permutationa(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 15, 16)]
        Double PDuration(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 15, 16)]
        string Base(Double arg1, Double arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        string Base(Double arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Decimal(string arg1, Double arg2);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Days(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Binom_Dist_Range(Double arg1, Double arg2, Double arg3, object arg4);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        Double Binom_Dist_Range(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Gamma(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Gauss(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Phi(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Rri(Double arg1, Double arg2, Double arg3);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        string Unichar(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Unicode(string arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        object Munit(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Arabic(string arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        Double IsoWeekNum(Double arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        Double IsoWeekNum(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        /// <param name="arg3">string arg3</param>
        [SupportByVersion("Excel", 15, 16)]
        Double NumberValue(string arg1, string arg2, string arg3);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        bool IsFormula(NetOffice.ExcelApi.Range arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        object IfNa(object arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Ceiling_Math(Double arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        Double Ceiling_Math(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        Double Ceiling_Math(Double arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 15, 16)]
        Double Floor_Math(Double arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        Double Floor_Math(Double arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        Double Floor_Math(Double arg1, object arg2);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        string ImSinh(object arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        string ImCosh(object arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        object FilterXML(string arg1, string arg2);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        object WebService(string arg1);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        object EncodeURL(string arg1);

        #endregion
    }
}
