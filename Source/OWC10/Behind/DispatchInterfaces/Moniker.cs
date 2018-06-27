using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api.Behind
{
    /// <summary>
    /// Moniker
    /// </summary>
    [SyntaxBypass]
    public class Moniker_ : COMObject, NetOffice.OWC10Api.Moniker_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Moniker_() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="relativeTo">optional object relativeTo</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_Moniker(object relativeTo)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Moniker", relativeTo);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Moniker
        /// </summary>
        /// <param name="relativeTo">optional object relativeTo</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Moniker")]
        public virtual string Moniker(object relativeTo)
        {
            return get_Moniker(relativeTo);
        }

        #endregion
    }

    /// <summary>
    /// DispatchInterface Moniker 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class Moniker : Moniker_, NetOffice.OWC10Api.Moniker
    {
        #pragma warning disable

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.OWC10Api.Moniker);
                return _contractType;
            }
        }
        private static Type _contractType;


        /// <summary>
        /// Instance Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type InstanceType
        {
            get
            {
                return LateBindingApiWrapperType;
            }
        }

        private static Type _type;

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Moniker);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public Moniker() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="moniker">string moniker</param>
        [SupportByVersion("OWC10", 1), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_Parse(string moniker)
        {
            return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parse", moniker);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Parse
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="moniker">string moniker</param>
        [SupportByVersion("OWC10", 1), ProxyResult, Redirect("get_Parse")]
        public virtual object Parse(string moniker)
        {
            return get_Parse(moniker);
        }

        #endregion

        #pragma warning restore
    }
}
