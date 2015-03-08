using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.ExcelApi.Tools
{
    /// <summary>
    /// Specify requested file format to get its extension in current application version
    /// </summary>
    public enum DocumentFormat
    {
        /// <summary>
        /// Default Workbook | xls or xlsx 
        /// </summary>
        Normal = 0,

        /// <summary>
        /// Workbook contains activated macros | xls or xlsm
        /// </summary>
        Macros = 1,

        /// <summary>
        /// Workbook Template | xlt or xltx
        /// </summary>
        Template = 2,

        /// <summary>
        /// Workbook Template contains activated macros | xlt or xltm
        /// </summary>
        TemplateMacros = 3,

        /// <summary>
        /// Binary Workbok, not support in 2003 or below | xlsb
        /// </summary>
        Binary = 4,

        /// <summary>
        /// Document Level Addin contains macros | xla or xlam
        /// </summary>
        AddinMacros = 5
    }
}
