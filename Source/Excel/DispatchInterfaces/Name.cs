using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// Name
    /// </summary>
    [SyntaxBypass]
    public interface Name_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840686.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Name { get; set; }

        #endregion
    }

    /// <summary>
    /// DispatchInterface Name 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839264.aspx </remarks>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface Name : Name_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840386.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835892.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834971.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string _Default { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838235.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Index { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192973.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Category { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821571.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string CategoryLocal { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197611.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlXLMMacroType MacroType { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834743.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object RefersTo { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841269.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string ShortcutKey { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195389.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Value { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193309.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool Visible { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834392.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string NameLocal { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840088.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object RefersToLocal { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822115.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object RefersToR1C1 { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196221.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object RefersToR1C1Local { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834918.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range RefersToRange { get; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834632.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        string Comment { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194620.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool WorkbookParameter { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841227.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool ValidWorkbookParameter { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197312.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void Delete();

        #endregion
    }
}
