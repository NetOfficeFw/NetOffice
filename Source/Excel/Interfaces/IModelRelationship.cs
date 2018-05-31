using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// Interface IModelRelationship 
    /// SupportByVersion Excel, 15, 16
    /// </summary>
    [SupportByVersion("Excel", 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public interface IModelRelationship : ICOMObject
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
        NetOffice.ExcelApi.ModelTable ForeignKeyTable { get; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.ModelTableColumn ForeignKeyColumn { get; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.ModelTable PrimaryKeyTable { get; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.ModelTableColumn PrimaryKeyColumn { get; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        bool Active { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        Int32 Delete();

        #endregion
    }
}
