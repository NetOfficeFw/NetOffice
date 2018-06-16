using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
    /// <summary>
    /// Name
    /// </summary>
    [SyntaxBypass]
    public interface Name_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        string Name { get; set; }

        #endregion
    }

    /// <summary>
    /// DispatchInterface Name 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("F5B39BAC-1480-11D3-8549-00C04FAC67D7")]
    public interface Name : Name_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        NetOffice.OWC10Api.ISpreadsheet Application { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        Int32 Index { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("OWC10", 1), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object RefersTo { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        object RefersToLocal { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        NetOffice.OWC10Api._Range RefersToRange { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        string Value { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        void Delete();

        #endregion
    }

   
}
