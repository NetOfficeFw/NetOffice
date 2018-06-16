using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
    /// <summary>
    /// ActiveObject
    /// </summary>
    [SyntaxBypass]
    public interface ActiveObject_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("OWC10", 1), ProxyResult]
        object ActiveObject { get; set; }

        #endregion
    }

    /// <summary>
    /// DispatchInterface ActiveObject 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("A809B678-545A-11D3-BE86-0050041DB15A")]
    public interface ActiveObject : ActiveObject_
    {

    } 
}
