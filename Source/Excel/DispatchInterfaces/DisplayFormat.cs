using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// DisplayFormat
    /// </summary>
    [SyntaxBypass]
    public interface DisplayFormat_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197441.aspx
        [SupportByVersion("Excel", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.ExcelApi.Characters get_Characters(object start, object length);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197441.aspx </remarks>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        [SupportByVersion("Excel", 14, 15, 16), Redirect("get_Characters")]
        NetOffice.ExcelApi.Characters Characters(object start, object length);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197441.aspx
        [SupportByVersion("Excel", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.ExcelApi.Characters get_Characters(object start);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197441.aspx </remarks>
        /// <param name="start">optional object start</param>
        [SupportByVersion("Excel", 14, 15, 16), Redirect("get_Characters")]
        NetOffice.ExcelApi.Characters Characters(object start);

        #endregion
    }

    /// <summary>
    /// DispatchInterface DisplayFormat 
    /// SupportByVersion Excel, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838863.aspx </remarks>
    [SupportByVersion("Excel", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("000244C2-0000-0000-C000-000000000046")]
    public interface DisplayFormat : DisplayFormat_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838383.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822645.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194522.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195472.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.Borders Borders { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197441.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        new NetOffice.ExcelApi.Characters Characters { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836734.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.Font Font { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196382.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        object Style { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821953.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        object AddIndent { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197010.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        object FormulaHidden { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821609.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        object HorizontalAlignment { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197906.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        object IndentLevel { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838619.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        NetOffice.ExcelApi.Interior Interior { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196267.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        object Locked { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837831.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        object MergeCells { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198350.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        object NumberFormat { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193870.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        object NumberFormatLocal { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840309.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        object Orientation { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839759.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        Int32 ReadingOrder { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834379.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        object ShrinkToFit { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837378.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        object VerticalAlignment { get; }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836493.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        object WrapText { get; }

        #endregion
    }
}
