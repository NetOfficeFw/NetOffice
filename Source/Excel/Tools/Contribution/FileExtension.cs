using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.ExcelApi.Tools.Contribution
{
    /// <summary>
    /// Represents well known excel file extensions
    /// </summary>
    public enum FileExtension
    {
        /// <summary>
        /// xlsx
        /// </summary>
        Workbook = 0,

        /// <summary>
        /// xls, may contains also macros
        /// </summary>
        WorkbookDepricated = 1,

        /// <summary>
        /// xlsm
        /// </summary>
        WorkbookInclMacros = 2,

        /// <summary>
        /// xltx
        /// </summary>
        Template = 2,

        /// <summary>
        /// xlt, may contains also macros
        /// </summary>
        TemplateDepricated = 3,

        /// <summary>
        /// xltm
        /// </summary>
        TemplateMacros = 4,

        /// <summary>
        /// xlsb
        /// </summary>
        Binary = 5,

        /// <summary>
        /// xla
        /// </summary>
        Addin = 6,

        /// <summary>
        /// xlam
        /// </summary>
        AddinMacros = 7,

        /// <summary>
        /// Unknown extension
        /// </summary>
        Unknown = 666
    }
}