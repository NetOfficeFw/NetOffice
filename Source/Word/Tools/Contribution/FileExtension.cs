using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.WordApi.Tools.Contribution
{
    /// <summary>
    ///  Represents well known word file extensions
    /// </summary>
    public enum FileExtension
    {
        /// <summary>
        /// docx
        /// </summary>
        Document = 0,

        /// <summary>
        /// doc, may include macros
        /// </summary>
        DocumentDepricated = 1,

        /// <summary>
        /// docxm
        /// </summary>
        DocumentInclMacros = 2,

        /// <summary>
        /// dotx
        /// </summary>
        Template = 3,

        /// <summary>
        /// dot, may include macros
        /// </summary>
        TemplateDepcricated = 4,

        /// <summary>
        /// dotm
        /// </summary>
        TemplateInclMacros = 5,

        /// <summary>
        /// Unknown extension
        /// </summary>
        Unknown = 666
    }
}