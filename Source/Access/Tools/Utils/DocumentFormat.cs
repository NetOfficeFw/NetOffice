using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.AccessApi.Tools
{
    /// <summary>
    /// Specify requested file format to get its extension in current application version
    /// </summary>
    public enum DocumentFormat
    {
        /// <summary>
        /// Default Database | mdb or accdb
        /// </summary>
        Normal = 0,

        /// <summary>
        /// Compiled database without source code | mde or accde
        /// </summary>
        Compiled = 1,

        /// <summary>
        /// Runtime database, not support in 2003 or below  | accdr
        /// </summary>
        Runtime = 2,

        /// <summary>
        /// Template database | mdt or accdt
        /// </summary>
        Template = 3,

        /// <summary>
        /// Addin database | mda or accda
        /// </summary>
        Addin = 4,

        /// <summary>
        /// Assistant database, not supported in 2007 or higher | mdz
        /// </summary>
        Assistant = 5,

        /// <summary>
        /// Field-Description Database, not supported in 2003 or below | accdf
        /// </summary>
        FieldDescription = 6,

        /// <summary>
        /// Workgroup Security File, not supported in 2007 or higher | mdw
        /// </summary>
        WorkgroupSecurity = 7
    }
}
