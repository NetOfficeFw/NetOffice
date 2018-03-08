using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.AccessApi.Tools.Contribution
{
    /// <summary>
    /// Represents well known access extensions
    /// </summary>
    public enum FileExtension
    {
        /// <summary>
        /// accdb
        /// </summary>
        Database = 0,

        /// <summary>
        /// mdb
        /// </summary>
        DatabaseDepricated = 1,

        /// <summary>
        /// accde
        /// </summary>
        CompiledDatabase = 2,

        /// <summary>
        /// mde
        /// </summary>
        CompiledDatabaseDepricated = 3,

        /// <summary>
        /// accdr
        /// </summary>
        RuntimeDatabase = 4,

        /// <summary>
        /// accdt
        /// </summary>
        Template = 6,

        /// <summary>
        /// mdt
        /// </summary>
        TemplateDepricated = 7,

        /// <summary>
        /// accda
        /// </summary>
        Addin = 8,

        /// <summary>
        /// mda
        /// </summary>
        AddinDepcricated = 9,

        /// <summary>
        /// mdz
        /// </summary>
        Assistant = 10,

        /// <summary>
        /// accdf
        /// </summary>
        FieldDescription = 11,

        /// <summary>
        /// mdw
        /// </summary>
        WorkgroupSecurity = 12,

        /// <summary>
        /// Unknown extension
        /// </summary>
        Unknown = 666
    }
}