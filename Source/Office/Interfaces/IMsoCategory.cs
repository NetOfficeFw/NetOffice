using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// Interface IMsoCategory 
    /// SupportByVersion Office, 15, 16
    /// </summary>
    [SupportByVersion("Office", 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public interface IMsoCategory : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 15,16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 15, 16)]
        string Name { get; }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 15, 16)]
        bool IsFiltered { get; set; }

        #endregion
    }
}
