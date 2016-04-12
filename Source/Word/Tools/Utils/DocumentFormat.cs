using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.WordApi.Tools
{
    /// <summary>
    /// Specify requested file format to get its extension in current application version
    /// </summary>
    public enum DocumentFormat
    {
        /// <summary>
        /// Default Document | doc or docx 
        /// </summary>
        Normal = 0,

        /// <summary>
        /// Document contains activated macros | doc or docm
        /// </summary>
        Macros = 1,

        /// <summary>
        /// Document Template | dot or dotx
        /// </summary>
        Template = 2,

        /// <summary>
        /// Document Template contains activated macros | dot or dotm
        /// </summary>
        TemplateMacros = 3

    }
}
