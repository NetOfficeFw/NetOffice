using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.PowerPointApi.Tools.Contribution
{
    /// <summary>
    /// Represents well known powerpoint extensions
    /// </summary>
    public enum FileExtension
    {
        /// <summary>
        /// pptx
        /// </summary>
        Presentation = 0,

        /// <summary>
        /// ppt, may include macros
        /// </summary>
        PresentationDepricated = 1,

        /// <summary>
        /// pptm
        /// </summary>
        PresentationInclMacros = 2,

        /// <summary>
        /// potx
        /// </summary>
        Template = 3,

        /// <summary>
        /// pot, may include macros
        /// </summary>
        TemplateDepcricated = 4,

        /// <summary>
        /// potm
        /// </summary>
        TemplateInclMacros = 5,

        /// <summary>
        /// ppsx
        /// </summary>
        RuntimePresentation = 6,

        /// <summary>
        /// pps, may include macros
        /// </summary>
        RuntimePresentationDepricated = 7,

        /// <summary>
        /// ppsm
        /// </summary>
        RuntimePresentationInclMacros = 8,

        /// <summary>
        /// Unknown extension
        /// </summary>
        Unknown = 666
    }
}