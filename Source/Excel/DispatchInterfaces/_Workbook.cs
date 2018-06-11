using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// _Workbook
    /// </summary>
    [SyntaxBypass]
    public interface _Workbook_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821660.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_Colors(object index);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_Colors(object index, object value);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Colors
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821660.aspx </remarks>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Colors")]
        object Colors(object index);

        #endregion
    }

    /// <summary>
    /// DispatchInterface _Workbook
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000208DA-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.ExcelApi.Workbook))]
    public interface _Workbook : _Workbook_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835918.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840080.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198008.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool AcceptLabelsInFormulas { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834923.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Chart ActiveChart { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841181.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object ActiveSheet { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string Author { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840067.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 AutoUpdateFrequency { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193298.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool AutoUpdateSaveChanges { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821530.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 ChangeHistoryDuration { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197172.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object BuiltinDocumentProperties { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821062.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Sheets Charts { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195162.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string CodeName { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string _CodeName { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821660.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        new object Colors { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835614.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBars CommandBars { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string Comments { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198339.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlSaveConflictResolution ConflictResolution { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834401.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object Container { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196337.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool CreateBackup { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834990.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object CustomDocumentProperties { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193264.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool Date1904 { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.ExcelApi.Sheets DialogSheets { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834329.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.xlDisplayDrawingObjects DisplayDrawingObjects { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840717.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlFileFormat FileFormat { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834975.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string FullName { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool HasMailer { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840238.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool HasPassword { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool HasRoutingSlip { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838249.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool IsAddin { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string Keywords { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837965.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Mailer Mailer { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.ExcelApi.Sheets Modules { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839882.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool MultiUserEditing { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820899.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Name { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195422.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Names Names { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string OnSave { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string OnSheetActivate { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string OnSheetDeactivate { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840974.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Path { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836500.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool PersonalViewListSettings { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822649.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool PersonalViewPrintSettings { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198189.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool PrecisionAsDisplayed { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838601.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool ProtectStructure { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193864.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool ProtectWindows { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840925.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool ReadOnly { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196964.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool ReadOnlyRecommended { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834665.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 RevisionNumber { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool Routed { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.RoutingSlip RoutingSlip { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196613.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool Saved { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840667.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool SaveLinkValues { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197568.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Sheets Sheets { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839677.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool ShowConflictHistory { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839039.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Styles Styles { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string Subject { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string Title { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193266.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool UpdateRemoteReferences { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool UserControl { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193788.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object UserStatus { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195531.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.CustomViews CustomViews { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195152.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Windows Windows { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835542.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Sheets Worksheets { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836228.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool WriteReserved { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840737.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string WriteReservedBy { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822819.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Sheets Excel4IntlMacroSheets { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195645.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Sheets Excel4MacroSheets { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836472.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool TemplateRemoveExtData { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194254.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool HighlightChangesOnScreen { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197016.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool KeepChangeHistory { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834301.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool ListChangesOnNewSheet { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194737.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.VBIDEApi.VBProject VBProject { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840963.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool IsInplace { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838208.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.PublishObjects PublishObjects { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834724.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.WebOptions WebOptions { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.HTMLProject HTMLProject { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839554.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool EnvelopeVisible { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193512.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 CalculationVersion { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822659.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool VBASigned { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool _ReadOnlyRecommended { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196322.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool ShowPivotTableFieldList { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839021.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlUpdateLinks UpdateLinks { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193225.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool EnableAutoRecover { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841017.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool RemovePersonalInformation { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821089.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        string FullNameURLEncoded { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821529.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        string Password { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837767.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        string WritePassword { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839579.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        string PasswordEncryptionProvider { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195464.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        string PasswordEncryptionAlgorithm { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195381.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        Int32 PasswordEncryptionKeyLength { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820819.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool PasswordEncryptionFileProperties { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.SmartTagOptions SmartTagOptions { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840697.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Permission Permission { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835236.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.SharedWorkspace SharedWorkspace { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192923.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Sync Sync { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838260.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.XmlNamespaces XmlNamespaces { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838975.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.XmlMaps XmlMaps { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194561.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.SmartDocument SmartDocument { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838205.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.DocumentLibraryVersions DocumentLibraryVersions { get; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837429.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        bool InactiveListBorderVisible { get; set; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838435.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        bool DisplayInkComments { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837152.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.OfficeApi.MetaProperties ContentTypeProperties { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836773.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.Connections Connections { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838073.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.OfficeApi.SignatureSet Signatures { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194489.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.OfficeApi.ServerPolicy ServerPolicy { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195426.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.OfficeApi.DocumentInspectors DocumentInspectors { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195818.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.ServerViewableItems ServerViewableItems { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837756.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.TableStyles TableStyles { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195934.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object DefaultTableStyle { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835624.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object DefaultPivotTableStyle { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836165.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool CheckCompatibility { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838063.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool HasVBProject { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838448.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLParts CustomXMLParts { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820907.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool Final { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196847.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.Research Research { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194072.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.OfficeApi.OfficeTheme Theme { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834991.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool Excel8CompatibilityMode { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837960.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool ConnectionsDisabled { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835280.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool ShowPivotChartActiveFields { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839003.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.IconSets IconSets { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194147.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string EncryptionProvider { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839440.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool DoNotPromptForConvert { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823189.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool ForceFullCalculation { get; set; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194925.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.SlicerCaches SlicerCaches { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839464.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.Slicer ActiveSlicer { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193862.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object DefaultSlicerStyle { get; set; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838425.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        Int32 AccuracyVersion { get; set; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229542.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        bool CaseSensitive { get; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231362.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        bool UseWholeCellCriteria { get; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230772.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        bool UseWildcards { get; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231277.aspx </remarks>
        [SupportByVersion("Excel", 15, 16), ProxyResult]
        object PivotTables { get; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228926.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        NetOffice.ExcelApi.Model Model { get; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227452.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        bool ChartDataPointTrack { get; set; }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230214.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object DefaultTimelineStyle { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821837.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Activate();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193344.aspx </remarks>
        /// <param name="mode">NetOffice.ExcelApi.Enums.XlFileAccess mode</param>
        /// <param name="writePassword">optional object writePassword</param>
        /// <param name="notify">optional object notify</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void ChangeFileAccess(NetOffice.ExcelApi.Enums.XlFileAccess mode, object writePassword, object notify);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193344.aspx </remarks>
        /// <param name="mode">NetOffice.ExcelApi.Enums.XlFileAccess mode</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void ChangeFileAccess(NetOffice.ExcelApi.Enums.XlFileAccess mode);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193344.aspx </remarks>
        /// <param name="mode">NetOffice.ExcelApi.Enums.XlFileAccess mode</param>
        /// <param name="writePassword">optional object writePassword</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void ChangeFileAccess(NetOffice.ExcelApi.Enums.XlFileAccess mode, object writePassword);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836537.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="newName">string newName</param>
        /// <param name="type">optional NetOffice.ExcelApi.Enums.XlLinkType Type = 1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void ChangeLink(string name, string newName, object type);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836537.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="newName">string newName</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void ChangeLink(string name, string newName);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838613.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="routeWorkbook">optional object routeWorkbook</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Close(object saveChanges, object filename, object routeWorkbook);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838613.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Close();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838613.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Close(object saveChanges);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838613.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        /// <param name="filename">optional object filename</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Close(object saveChanges, object filename);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839565.aspx </remarks>
        /// <param name="numberFormat">string numberFormat</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void DeleteNumberFormat(string numberFormat);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836762.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool ExclusiveAccess();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836208.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void ForwardMailer();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192971.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="linkInfo">NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo</param>
        /// <param name="type">optional object type</param>
        /// <param name="editionRef">optional object editionRef</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object LinkInfo(string name, NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo, object type, object editionRef);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192971.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="linkInfo">NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object LinkInfo(string name, NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192971.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="linkInfo">NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo</param>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object LinkInfo(string name, NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo, object type);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821922.aspx </remarks>
        /// <param name="type">optional object type</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object LinkSources(object type);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821922.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object LinkSources();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196693.aspx </remarks>
        /// <param name="filename">object filename</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void MergeWorkbook(object filename);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838378.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Window NewWindow();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839052.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="readOnly">optional object readOnly</param>
        /// <param name="type">optional object type</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void OpenLinks(string name, object readOnly, object type);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839052.aspx </remarks>
        /// <param name="name">string name</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void OpenLinks(string name);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839052.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="readOnly">optional object readOnly</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void OpenLinks(string name, object readOnly);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193549.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.PivotCaches PivotCaches();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821844.aspx </remarks>
        /// <param name="destName">optional object destName</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Post(object destName);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821844.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Post();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="prToFileName">optional object prToFileName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void _PrintOut();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void _PrintOut(object from);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void _PrintOut(object from, object to);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void _PrintOut(object from, object to, object copies);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void _PrintOut(object from, object to, object copies, object preview);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void _PrintOut(object from, object to, object copies, object preview, object activePrinter);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193068.aspx </remarks>
        /// <param name="enableChanges">optional object enableChanges</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PrintPreview(object enableChanges);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193068.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PrintPreview();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193800.aspx </remarks>
        /// <param name="password">optional object password</param>
        /// <param name="structure">optional object structure</param>
        /// <param name="windows">optional object windows</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Protect(object password, object structure, object windows);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193800.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Protect();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193800.aspx </remarks>
        /// <param name="password">optional object password</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Protect(object password);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193800.aspx </remarks>
        /// <param name="password">optional object password</param>
        /// <param name="structure">optional object structure</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Protect(object password, object structure);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="sharingPassword">optional object sharingPassword</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object sharingPassword);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="sharingPassword">optional object sharingPassword</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object sharingPassword, object fileFormat);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void ProtectSharing();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void ProtectSharing(object filename);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void ProtectSharing(object filename, object password);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void ProtectSharing(object filename, object password, object writeResPassword);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838648.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void RefreshAll();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820902.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Reply();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838788.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void ReplyAll();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840747.aspx </remarks>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void RemoveUser(Int32 index);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Route();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835203.aspx </remarks>
        /// <param name="which">NetOffice.ExcelApi.Enums.XlRunAutoMacro which</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void RunAutoMacros(NetOffice.ExcelApi.Enums.XlRunAutoMacro which);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197585.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Save();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        /// <param name="conflictResolution">optional object conflictResolution</param>
        /// <param name="addToMru">optional object addToMru</param>
        /// <param name="textCodepage">optional object textCodepage</param>
        /// <param name="textVisualLayout">optional object textVisualLayout</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage, object textVisualLayout);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        /// <param name="conflictResolution">optional object conflictResolution</param>
        /// <param name="addToMru">optional object addToMru</param>
        /// <param name="textCodepage">optional object textCodepage</param>
        /// <param name="textVisualLayout">optional object textVisualLayout</param>
        /// <param name="local">optional object local</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage, object textVisualLayout, object local);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SaveAs();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SaveAs(object filename);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SaveAs(object filename, object fileFormat);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SaveAs(object filename, object fileFormat, object password);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SaveAs(object filename, object fileFormat, object password, object writeResPassword);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        /// <param name="conflictResolution">optional object conflictResolution</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        /// <param name="conflictResolution">optional object conflictResolution</param>
        /// <param name="addToMru">optional object addToMru</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        /// <param name="conflictResolution">optional object conflictResolution</param>
        /// <param name="addToMru">optional object addToMru</param>
        /// <param name="textCodepage">optional object textCodepage</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835014.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SaveCopyAs(object filename);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835014.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SaveCopyAs();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821053.aspx </remarks>
        /// <param name="recipients">object recipients</param>
        /// <param name="subject">optional object subject</param>
        /// <param name="returnReceipt">optional object returnReceipt</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SendMail(object recipients, object subject, object returnReceipt);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821053.aspx </remarks>
        /// <param name="recipients">object recipients</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SendMail(object recipients);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821053.aspx </remarks>
        /// <param name="recipients">object recipients</param>
        /// <param name="subject">optional object subject</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SendMail(object recipients, object subject);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840261.aspx </remarks>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="priority">optional NetOffice.ExcelApi.Enums.XlPriority Priority = -4143</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SendMailer(object fileFormat, object priority);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840261.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SendMailer();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840261.aspx </remarks>
        /// <param name="fileFormat">optional object fileFormat</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SendMailer(object fileFormat);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838177.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="procedure">optional object procedure</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SetLinkOnData(string name, object procedure);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838177.aspx </remarks>
        /// <param name="name">string name</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void SetLinkOnData(string name);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196695.aspx </remarks>
        /// <param name="password">optional object password</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Unprotect(object password);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196695.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Unprotect();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840639.aspx </remarks>
        /// <param name="sharingPassword">optional object sharingPassword</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void UnprotectSharing(object sharingPassword);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840639.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void UnprotectSharing();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840979.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void UpdateFromFile();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195741.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="type">optional object type</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void UpdateLink(object name, object type);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195741.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void UpdateLink();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195741.aspx </remarks>
        /// <param name="name">optional object name</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void UpdateLink(object name);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837763.aspx </remarks>
        /// <param name="when">optional object when</param>
        /// <param name="who">optional object who</param>
        /// <param name="where">optional object where</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void HighlightChangesOptions(object when, object who, object where);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837763.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void HighlightChangesOptions();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837763.aspx </remarks>
        /// <param name="when">optional object when</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void HighlightChangesOptions(object when);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837763.aspx </remarks>
        /// <param name="when">optional object when</param>
        /// <param name="who">optional object who</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void HighlightChangesOptions(object when, object who);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834662.aspx </remarks>
        /// <param name="days">Int32 days</param>
        /// <param name="sharingPassword">optional object sharingPassword</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PurgeChangeHistoryNow(Int32 days, object sharingPassword);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834662.aspx </remarks>
        /// <param name="days">Int32 days</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PurgeChangeHistoryNow(Int32 days);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835613.aspx </remarks>
        /// <param name="when">optional object when</param>
        /// <param name="who">optional object who</param>
        /// <param name="where">optional object where</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void AcceptAllChanges(object when, object who, object where);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835613.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void AcceptAllChanges();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835613.aspx </remarks>
        /// <param name="when">optional object when</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void AcceptAllChanges(object when);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835613.aspx </remarks>
        /// <param name="when">optional object when</param>
        /// <param name="who">optional object who</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void AcceptAllChanges(object when, object who);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837407.aspx </remarks>
        /// <param name="when">optional object when</param>
        /// <param name="who">optional object who</param>
        /// <param name="where">optional object where</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void RejectAllChanges(object when, object who, object where);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837407.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void RejectAllChanges();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837407.aspx </remarks>
        /// <param name="when">optional object when</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void RejectAllChanges(object when);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837407.aspx </remarks>
        /// <param name="when">optional object when</param>
        /// <param name="who">optional object who</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void RejectAllChanges(object when, object who);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        /// <param name="backgroundQuery">optional object backgroundQuery</param>
        /// <param name="optimizeCache">optional object optimizeCache</param>
        /// <param name="pageFieldOrder">optional object pageFieldOrder</param>
        /// <param name="pageFieldWrapCount">optional object pageFieldWrapCount</param>
        /// <param name="readData">optional object readData</param>
        /// <param name="connection">optional object connection</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount, object readData, object connection);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        /// <param name="backgroundQuery">optional object backgroundQuery</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        /// <param name="backgroundQuery">optional object backgroundQuery</param>
        /// <param name="optimizeCache">optional object optimizeCache</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        /// <param name="backgroundQuery">optional object backgroundQuery</param>
        /// <param name="optimizeCache">optional object optimizeCache</param>
        /// <param name="pageFieldOrder">optional object pageFieldOrder</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        /// <param name="backgroundQuery">optional object backgroundQuery</param>
        /// <param name="optimizeCache">optional object optimizeCache</param>
        /// <param name="pageFieldOrder">optional object pageFieldOrder</param>
        /// <param name="pageFieldWrapCount">optional object pageFieldWrapCount</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        /// <param name="backgroundQuery">optional object backgroundQuery</param>
        /// <param name="optimizeCache">optional object optimizeCache</param>
        /// <param name="pageFieldOrder">optional object pageFieldOrder</param>
        /// <param name="pageFieldWrapCount">optional object pageFieldWrapCount</param>
        /// <param name="readData">optional object readData</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount, object readData);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194697.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void ResetColors();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
        /// <param name="address">string address</param>
        /// <param name="subAddress">optional object subAddress</param>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="addHistory">optional object addHistory</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        /// <param name="method">optional object method</param>
        /// <param name="headerInfo">optional object headerInfo</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method, object headerInfo);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
        /// <param name="address">string address</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void FollowHyperlink(string address);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
        /// <param name="address">string address</param>
        /// <param name="subAddress">optional object subAddress</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void FollowHyperlink(string address, object subAddress);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
        /// <param name="address">string address</param>
        /// <param name="subAddress">optional object subAddress</param>
        /// <param name="newWindow">optional object newWindow</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void FollowHyperlink(string address, object subAddress, object newWindow);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
        /// <param name="address">string address</param>
        /// <param name="subAddress">optional object subAddress</param>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="addHistory">optional object addHistory</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
        /// <param name="address">string address</param>
        /// <param name="subAddress">optional object subAddress</param>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="addHistory">optional object addHistory</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
        /// <param name="address">string address</param>
        /// <param name="subAddress">optional object subAddress</param>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="addHistory">optional object addHistory</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        /// <param name="method">optional object method</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194282.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void AddToFavorites();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="prToFileName">optional object prToFileName</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="prToFileName">optional object prToFileName</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName, object ignorePrintAreas);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PrintOut();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        /// <param name="from">optional object from</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PrintOut(object from);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PrintOut(object from, object to);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PrintOut(object from, object to, object copies);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PrintOut(object from, object to, object copies, object preview);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PrintOut(object from, object to, object copies, object preview, object activePrinter);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195831.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void WebPagePreview();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839234.aspx </remarks>
        /// <param name="encoding">NetOffice.OfficeApi.Enums.MsoEncoding encoding</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void ReloadAs(NetOffice.OfficeApi.Enums.MsoEncoding encoding);

        /// <summary>
        /// SupportByVersion Excel 9
        /// </summary>
        /// <param name="unused">Int32 unused</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9)]
        void Dummy1(Int32 unused);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="s">string s</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void sblt(string s);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="structure">optional object structure</param>
        /// <param name="windows">optional object windows</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void _Protect(object password, object structure, object windows);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void _Protect();

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void _Protect(object password);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="structure">optional object structure</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void _Protect(object password, object structure);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        /// <param name="conflictResolution">optional object conflictResolution</param>
        /// <param name="addToMru">optional object addToMru</param>
        /// <param name="textCodepage">optional object textCodepage</param>
        /// <param name="textVisualLayout">optional object textVisualLayout</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage, object textVisualLayout);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void _SaveAs();

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void _SaveAs(object filename);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void _SaveAs(object filename, object fileFormat);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void _SaveAs(object filename, object fileFormat, object password);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void _SaveAs(object filename, object fileFormat, object password, object writeResPassword);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        /// <param name="conflictResolution">optional object conflictResolution</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        /// <param name="conflictResolution">optional object conflictResolution</param>
        /// <param name="addToMru">optional object addToMru</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        /// <param name="conflictResolution">optional object conflictResolution</param>
        /// <param name="addToMru">optional object addToMru</param>
        /// <param name="textCodepage">optional object textCodepage</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="calcid">Int32 calcid</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void Dummy17(Int32 calcid);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194915.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlLinkType type</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void BreakLink(string name, NetOffice.ExcelApi.Enums.XlLinkType type);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void Dummy16();

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841145.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        /// <param name="comments">optional object comments</param>
        /// <param name="makePublic">optional object makePublic</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void CheckIn(object saveChanges, object comments, object makePublic);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841145.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void CheckIn();

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841145.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void CheckIn(object saveChanges);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841145.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        /// <param name="comments">optional object comments</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void CheckIn(object saveChanges, object comments);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194456.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool CanCheckIn();

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx </remarks>
        /// <param name="recipients">optional object recipients</param>
        /// <param name="subject">optional object subject</param>
        /// <param name="showMessage">optional object showMessage</param>
        /// <param name="includeAttachment">optional object includeAttachment</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void SendForReview(object recipients, object subject, object showMessage, object includeAttachment);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void SendForReview();

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx </remarks>
        /// <param name="recipients">optional object recipients</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void SendForReview(object recipients);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx </remarks>
        /// <param name="recipients">optional object recipients</param>
        /// <param name="subject">optional object subject</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void SendForReview(object recipients, object subject);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx </remarks>
        /// <param name="recipients">optional object recipients</param>
        /// <param name="subject">optional object subject</param>
        /// <param name="showMessage">optional object showMessage</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void SendForReview(object recipients, object subject, object showMessage);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821626.aspx </remarks>
        /// <param name="showMessage">optional object showMessage</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void ReplyWithChanges(object showMessage);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821626.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void ReplyWithChanges();

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839207.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void EndReview();

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx </remarks>
        /// <param name="passwordEncryptionProvider">optional object passwordEncryptionProvider</param>
        /// <param name="passwordEncryptionAlgorithm">optional object passwordEncryptionAlgorithm</param>
        /// <param name="passwordEncryptionKeyLength">optional object passwordEncryptionKeyLength</param>
        /// <param name="passwordEncryptionFileProperties">optional object passwordEncryptionFileProperties</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void SetPasswordEncryptionOptions(object passwordEncryptionProvider, object passwordEncryptionAlgorithm, object passwordEncryptionKeyLength, object passwordEncryptionFileProperties);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void SetPasswordEncryptionOptions();

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx </remarks>
        /// <param name="passwordEncryptionProvider">optional object passwordEncryptionProvider</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void SetPasswordEncryptionOptions(object passwordEncryptionProvider);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx </remarks>
        /// <param name="passwordEncryptionProvider">optional object passwordEncryptionProvider</param>
        /// <param name="passwordEncryptionAlgorithm">optional object passwordEncryptionAlgorithm</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void SetPasswordEncryptionOptions(object passwordEncryptionProvider, object passwordEncryptionAlgorithm);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx </remarks>
        /// <param name="passwordEncryptionProvider">optional object passwordEncryptionProvider</param>
        /// <param name="passwordEncryptionAlgorithm">optional object passwordEncryptionAlgorithm</param>
        /// <param name="passwordEncryptionKeyLength">optional object passwordEncryptionKeyLength</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void SetPasswordEncryptionOptions(object passwordEncryptionProvider, object passwordEncryptionAlgorithm, object passwordEncryptionKeyLength);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        void RecheckSmartTags();

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840370.aspx </remarks>
        /// <param name="recipients">optional object recipients</param>
        /// <param name="subject">optional object subject</param>
        /// <param name="showMessage">optional object showMessage</param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        void SendFaxOverInternet(object recipients, object subject, object showMessage);

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840370.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        void SendFaxOverInternet();

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840370.aspx </remarks>
        /// <param name="recipients">optional object recipients</param>
        [CustomMethod]
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        void SendFaxOverInternet(object recipients);

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840370.aspx </remarks>
        /// <param name="recipients">optional object recipients</param>
        /// <param name="subject">optional object subject</param>
        [CustomMethod]
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        void SendFaxOverInternet(object recipients, object subject);

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836452.aspx </remarks>
        /// <param name="url">string url</param>
        /// <param name="importMap">NetOffice.ExcelApi.XmlMap importMap</param>
        /// <param name="overwrite">optional object overwrite</param>
        /// <param name="destination">optional object destination</param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImport(string url, out NetOffice.ExcelApi.XmlMap importMap, object overwrite, object destination);

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836452.aspx </remarks>
        /// <param name="url">string url</param>
        /// <param name="importMap">NetOffice.ExcelApi.XmlMap importMap</param>
        [CustomMethod]
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImport(string url, out NetOffice.ExcelApi.XmlMap importMap);

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836452.aspx </remarks>
        /// <param name="url">string url</param>
        /// <param name="importMap">NetOffice.ExcelApi.XmlMap importMap</param>
        /// <param name="overwrite">optional object overwrite</param>
        [CustomMethod]
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImport(string url, out NetOffice.ExcelApi.XmlMap importMap, object overwrite);

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837991.aspx </remarks>
        /// <param name="data">string data</param>
        /// <param name="importMap">NetOffice.ExcelApi.XmlMap importMap</param>
        /// <param name="overwrite">optional object overwrite</param>
        /// <param name="destination">optional object destination</param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImportXml(string data, out NetOffice.ExcelApi.XmlMap importMap, object overwrite, object destination);

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837991.aspx </remarks>
        /// <param name="data">string data</param>
        /// <param name="importMap">NetOffice.ExcelApi.XmlMap importMap</param>
        [CustomMethod]
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImportXml(string data, out NetOffice.ExcelApi.XmlMap importMap);

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837991.aspx </remarks>
        /// <param name="data">string data</param>
        /// <param name="importMap">NetOffice.ExcelApi.XmlMap importMap</param>
        /// <param name="overwrite">optional object overwrite</param>
        [CustomMethod]
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImportXml(string data, out NetOffice.ExcelApi.XmlMap importMap, object overwrite);

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834616.aspx </remarks>
        /// <param name="filename">string filename</param>
        /// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        void SaveAsXMLData(string filename, NetOffice.ExcelApi.XmlMap map);

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196845.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        void ToggleFormsDesign();

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void __PrintOut();

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void __PrintOut(object from);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void __PrintOut(object from, object to);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void __PrintOut(object from, object to, object copies);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void __PrintOut(object from, object to, object copies, object preview);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void __PrintOut(object from, object to, object copies, object preview, object activePrinter);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="sharingPassword">optional object sharingPassword</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void _ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object sharingPassword);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void _ProtectSharing();

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void _ProtectSharing(object filename);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void _ProtectSharing(object filename, object password);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void _ProtectSharing(object filename, object password, object writeResPassword);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void _ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void _ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840327.aspx </remarks>
        /// <param name="removeDocInfoType">NetOffice.ExcelApi.Enums.XlRemoveDocInfoType removeDocInfoType</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void RemoveDocumentInformation(NetOffice.ExcelApi.Enums.XlRemoveDocInfoType removeDocInfoType);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        /// <param name="comments">optional object comments</param>
        /// <param name="makePublic">optional object makePublic</param>
        /// <param name="versionType">optional object versionType</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void CheckInWithVersion(object saveChanges, object comments, object makePublic, object versionType);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void CheckInWithVersion();

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void CheckInWithVersion(object saveChanges);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        /// <param name="comments">optional object comments</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void CheckInWithVersion(object saveChanges, object comments);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        /// <param name="comments">optional object comments</param>
        /// <param name="makePublic">optional object makePublic</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void CheckInWithVersion(object saveChanges, object comments, object makePublic);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838567.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void LockServerFile();

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835507.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.OfficeApi.WorkflowTasks GetWorkflowTasks();

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837818.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.OfficeApi.WorkflowTemplates GetWorkflowTemplates();

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194014.aspx </remarks>
        /// <param name="filename">string filename</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ApplyTheme(string filename);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820742.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void EnableConnections();

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="openAfterPublish">optional object openAfterPublish</param>
        /// <param name="fixedFormatExtClassPtr">optional object fixedFormatExtClassPtr</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish, object fixedFormatExtClassPtr);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        /// <param name="from">optional object from</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to);

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="openAfterPublish">optional object openAfterPublish</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 14, 15, 16)]
        void Dummy26();

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 14, 15, 16)]
        void Dummy27();

        #endregion
    }
}
