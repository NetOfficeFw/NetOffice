using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// PageSetup
    /// </summary>
    [SyntaxBypass]
    public interface PageSetup_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194768.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_PrintQuality(object index);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_PrintQuality(object index, object value);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_PrintQuality
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194768.aspx </remarks>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_PrintQuality")]
        object PrintQuality(object index);

        #endregion
    }

    /// <summary>
    /// DispatchInterface PageSetup 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196103.aspx </remarks>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface PageSetup : PageSetup_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836795.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835290.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836825.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834921.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool BlackAndWhite { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198024.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double BottomMargin { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839224.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string CenterFooter { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835588.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string CenterHeader { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822498.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool CenterHorizontally { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841275.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool CenterVertically { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlObjectSize ChartSize { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194155.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool Draft { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821636.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 FirstPageNumber { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194582.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object FitToPagesTall { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194365.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object FitToPagesWide { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838218.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double FooterMargin { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838659.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double HeaderMargin { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822322.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string LeftFooter { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836805.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string LeftHeader { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821496.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double LeftMargin { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839060.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlOrder Order { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836834.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlPageOrientation Orientation { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834612.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlPaperSize PaperSize { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839804.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string PrintArea { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836524.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool PrintGridlines { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192995.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool PrintHeadings { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822102.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool PrintNotes { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194768.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object PrintQuality { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835214.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string PrintTitleColumns { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839961.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string PrintTitleRows { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840096.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string RightFooter { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836463.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string RightHeader { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836771.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double RightMargin { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194770.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Double TopMargin { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197028.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object Zoom { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194941.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlPrintLocation PrintComments { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197865.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlPrintErrors PrintErrors { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838811.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Graphic CenterHeaderPicture { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822635.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Graphic CenterFooterPicture { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194861.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Graphic LeftHeaderPicture { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195659.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Graphic LeftFooterPicture { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196670.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Graphic RightHeaderPicture { get; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840868.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Graphic RightFooterPicture { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839665.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool OddAndEvenPagesHeaderFooter { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838040.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool DifferentFirstPageHeaderFooter { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838420.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool ScaleWithDocHeaderFooter { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838197.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool AlignMarginsHeaderFooter { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840030.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.Pages Pages { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839989.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.Page EvenPage { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835921.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        NetOffice.ExcelApi.Page FirstPage { get; }

        #endregion
    }
}
