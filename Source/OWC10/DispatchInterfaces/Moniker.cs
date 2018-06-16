using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
    /// <summary>
    /// Moniker
    /// </summary>
    [SyntaxBypass]
    public interface Moniker_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="relativeTo">optional object relativeTo</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_Moniker(object relativeTo);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Moniker
        /// </summary>
        /// <param name="relativeTo">optional object relativeTo</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Moniker")]
        string Moniker(object relativeTo);

        #endregion
    }

    /// <summary>
    /// DispatchInterface Moniker 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("5055F752-6848-4CEA-9BAB-265EC4B5380A")]
    public interface Moniker : Moniker_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="moniker">string moniker</param>
        [SupportByVersion("OWC10", 1), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_Parse(string moniker);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Parse
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="moniker">string moniker</param>
        [SupportByVersion("OWC10", 1), ProxyResult, Redirect("get_Parse")]
        object Parse(string moniker);

        #endregion
    }
}
