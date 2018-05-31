using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// Interface IModel 
    /// SupportByVersion Excel, 15, 16
    /// </summary>
    [SupportByVersion("Excel", 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public interface IModel : ICOMObject
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
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.ModelTables ModelTables { get; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.ModelRelationships ModelRelationships { get; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.WorkbookConnection DataModelConnection { get; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        string Name { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        Int32 Refresh();

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="connectionToDataSource">NetOffice.ExcelApi.WorkbookConnection connectionToDataSource</param>
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.WorkbookConnection AddConnection(NetOffice.ExcelApi.WorkbookConnection connectionToDataSource);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="modelTable">object modelTable</param>
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.WorkbookConnection CreateModelWorkbookConnection(object modelTable);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        Int32 Initialize();

        #endregion
    }
}
