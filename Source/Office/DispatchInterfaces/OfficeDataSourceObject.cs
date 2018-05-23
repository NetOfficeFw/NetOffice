using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// DispatchInterface OfficeDataSourceObject 
    /// SupportByVersion Office, 10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864883.aspx </remarks>
    [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface OfficeDataSourceObject : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861793.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string ConnectString { get; set; }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861897.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string Table { get; set; }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860869.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        string DataSource { get; set; }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860229.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16), ProxyResult]
        object Columns { get; }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861767.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 RowCount { get; }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860598.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16), ProxyResult]
        object Filters { get;}

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864664.aspx </remarks>
        /// <param name="msoMoveRow">NetOffice.OfficeApi.Enums.MsoMoveRow msoMoveRow</param>
        /// <param name="rowNbr">optional Int32 RowNbr = 1</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 Move(NetOffice.OfficeApi.Enums.MsoMoveRow msoMoveRow, object rowNbr);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864664.aspx </remarks>
        /// <param name="msoMoveRow">NetOffice.OfficeApi.Enums.MsoMoveRow msoMoveRow</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 Move(NetOffice.OfficeApi.Enums.MsoMoveRow msoMoveRow);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865289.aspx </remarks>
        /// <param name="bstrSrc">optional string bstrSrc = </param>
        /// <param name="bstrConnect">optional string bstrConnect = </param>
        /// <param name="bstrTable">optional string bstrTable = </param>
        /// <param name="fOpenExclusive">optional Int32 fOpenExclusive = 0</param>
        /// <param name="fNeverPrompt">optional Int32 fNeverPrompt = 1</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void Open(object bstrSrc, object bstrConnect, object bstrTable, object fOpenExclusive, object fNeverPrompt);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865289.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void Open();

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865289.aspx </remarks>
        /// <param name="bstrSrc">optional string bstrSrc = </param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void Open(object bstrSrc);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865289.aspx </remarks>
        /// <param name="bstrSrc">optional string bstrSrc = </param>
        /// <param name="bstrConnect">optional string bstrConnect = </param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void Open(object bstrSrc, object bstrConnect);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865289.aspx </remarks>
        /// <param name="bstrSrc">optional string bstrSrc = </param>
        /// <param name="bstrConnect">optional string bstrConnect = </param>
        /// <param name="bstrTable">optional string bstrTable = </param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void Open(object bstrSrc, object bstrConnect, object bstrTable);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865289.aspx </remarks>
        /// <param name="bstrSrc">optional string bstrSrc = </param>
        /// <param name="bstrConnect">optional string bstrConnect = </param>
        /// <param name="bstrTable">optional string bstrTable = </param>
        /// <param name="fOpenExclusive">optional Int32 fOpenExclusive = 0</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void Open(object bstrSrc, object bstrConnect, object bstrTable, object fOpenExclusive);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861392.aspx </remarks>
        /// <param name="sortField1">string sortField1</param>
        /// <param name="sortAscending1">optional bool SortAscending1 = true</param>
        /// <param name="sortField2">optional string SortField2 = </param>
        /// <param name="sortAscending2">optional bool SortAscending2 = true</param>
        /// <param name="sortField3">optional string SortField3 = </param>
        /// <param name="sortAscending3">optional bool SortAscending3 = true</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetSortOrder(string sortField1, object sortAscending1, object sortField2, object sortAscending2, object sortField3, object sortAscending3);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861392.aspx </remarks>
        /// <param name="sortField1">string sortField1</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetSortOrder(string sortField1);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861392.aspx </remarks>
        /// <param name="sortField1">string sortField1</param>
        /// <param name="sortAscending1">optional bool SortAscending1 = true</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetSortOrder(string sortField1, object sortAscending1);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861392.aspx </remarks>
        /// <param name="sortField1">string sortField1</param>
        /// <param name="sortAscending1">optional bool SortAscending1 = true</param>
        /// <param name="sortField2">optional string SortField2 = </param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetSortOrder(string sortField1, object sortAscending1, object sortField2);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861392.aspx </remarks>
        /// <param name="sortField1">string sortField1</param>
        /// <param name="sortAscending1">optional bool SortAscending1 = true</param>
        /// <param name="sortField2">optional string SortField2 = </param>
        /// <param name="sortAscending2">optional bool SortAscending2 = true</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetSortOrder(string sortField1, object sortAscending1, object sortField2, object sortAscending2);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861392.aspx </remarks>
        /// <param name="sortField1">string sortField1</param>
        /// <param name="sortAscending1">optional bool SortAscending1 = true</param>
        /// <param name="sortField2">optional string SortField2 = </param>
        /// <param name="sortAscending2">optional bool SortAscending2 = true</param>
        /// <param name="sortField3">optional string SortField3 = </param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void SetSortOrder(string sortField1, object sortAscending1, object sortField2, object sortAscending2, object sortField3);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863341.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void ApplyFilter();

        #endregion
    }
}
