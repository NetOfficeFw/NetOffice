using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.PowerPointApi.Tools
{
    /// <summary>
    /// Specify requested file format to get its extension in current application version
    /// </summary>
    public enum DocumentFormat
    {
        /// <summary>
        /// Default Document | ppt or pptx 
        /// </summary>
        Normal = 0,

        /// <summary>
        /// Document contains activated macros | ppt or pptm
        /// </summary>
        Macros = 1,

        /// <summary>
        /// Document Template | pot or potx
        /// </summary>
        Template = 2,

        /// <summary>
        /// Document Template contains activated macros | pot or potm
        /// </summary>
        TemplateMacros = 3,

        /// <summary>
        /// Runtime Presentation | pps or ppsx
        /// </summary>
        Presentation = 4,

        /// <summary>
        /// Runtime Presentation contains activated macros | pps or ppsm
        /// </summary>
        PresentationMacros = 5
    }
}
