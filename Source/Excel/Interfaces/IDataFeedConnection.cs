using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// Interface IDataFeedConnection 
    /// SupportByVersion Excel, 15, 16
    /// </summary>
    [SupportByVersion("Excel", 15, 16)]
    [EntityType(EntityType.IsInterface)]
	[TypeId("000244D4-0001-0000-C000-000000000046")]
    public interface IDataFeedConnection : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        bool AlwaysUseConnectionFile { get; set; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        object CommandText { get; set; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.Enums.XlCmdType CommandType { get; set; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        object Connection { get; set; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        bool EnableRefresh { get; set; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        DateTime RefreshDate { get;}

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        bool Refreshing { get; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        bool RefreshOnFileOpen { get; set; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        Int32 RefreshPeriod { get; set; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        bool SavePassword { get; set; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.Enums.XlCredentialsMethod ServerCredentialsMethod { get; set; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        string SourceConnectionFile { get; set; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        string SourceDataFile { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        Int32 CancelRefresh();

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        Int32 Refresh();

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="oDCFileName">string oDCFileName</param>
        /// <param name="description">optional object description</param>
        /// <param name="keywords">optional object keywords</param>
        [SupportByVersion("Excel", 15, 16)]
        Int32 SaveAsODC(string oDCFileName, object description, object keywords);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="oDCFileName">string oDCFileName</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        Int32 SaveAsODC(string oDCFileName);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="oDCFileName">string oDCFileName</param>
        /// <param name="description">optional object description</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        Int32 SaveAsODC(string oDCFileName, object description);

        #endregion
    }
}
