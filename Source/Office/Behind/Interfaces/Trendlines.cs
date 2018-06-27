using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// Interface Trendlines 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    public class Trendlines : COMObject, NetOffice.OfficeApi.Trendlines
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
                    _contractType = typeof(NetOffice.OfficeApi.Trendlines);
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
                    _type = typeof(Trendlines);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Trendlines() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 Count
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16), ProxyResult]
        public virtual object Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Application");
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Int32 Creator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Office", 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public virtual NetOffice.OfficeApi.IMsoTrendline this[object index]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoTrendline>(this, "_Default", typeof(NetOffice.OfficeApi.IMsoTrendline), index);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
        /// <param name="order">optional object order</param>
        /// <param name="period">optional object period</param>
        /// <param name="forward">optional object forward</param>
        /// <param name="backward">optional object backward</param>
        /// <param name="intercept">optional object intercept</param>
        /// <param name="displayEquation">optional object displayEquation</param>
        /// <param name="displayRSquared">optional object displayRSquared</param>
        /// <param name="name">optional object name</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation, object displayRSquared, object name)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", typeof(NetOffice.OfficeApi.IMsoTrendline), new object[] { type, order, period, forward, backward, intercept, displayEquation, displayRSquared, name });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoTrendline Add()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", typeof(NetOffice.OfficeApi.IMsoTrendline));
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoTrendline Add(object type)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", typeof(NetOffice.OfficeApi.IMsoTrendline), type);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
        /// <param name="order">optional object order</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoTrendline Add(object type, object order)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", typeof(NetOffice.OfficeApi.IMsoTrendline), type, order);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
        /// <param name="order">optional object order</param>
        /// <param name="period">optional object period</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", typeof(NetOffice.OfficeApi.IMsoTrendline), type, order, period);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
        /// <param name="order">optional object order</param>
        /// <param name="period">optional object period</param>
        /// <param name="forward">optional object forward</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period, object forward)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", typeof(NetOffice.OfficeApi.IMsoTrendline), type, order, period, forward);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
        /// <param name="order">optional object order</param>
        /// <param name="period">optional object period</param>
        /// <param name="forward">optional object forward</param>
        /// <param name="backward">optional object backward</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period, object forward, object backward)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", typeof(NetOffice.OfficeApi.IMsoTrendline), new object[] { type, order, period, forward, backward });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
        /// <param name="order">optional object order</param>
        /// <param name="period">optional object period</param>
        /// <param name="forward">optional object forward</param>
        /// <param name="backward">optional object backward</param>
        /// <param name="intercept">optional object intercept</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period, object forward, object backward, object intercept)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", typeof(NetOffice.OfficeApi.IMsoTrendline), new object[] { type, order, period, forward, backward, intercept });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
        /// <param name="order">optional object order</param>
        /// <param name="period">optional object period</param>
        /// <param name="forward">optional object forward</param>
        /// <param name="backward">optional object backward</param>
        /// <param name="intercept">optional object intercept</param>
        /// <param name="displayEquation">optional object displayEquation</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", typeof(NetOffice.OfficeApi.IMsoTrendline), new object[] { type, order, period, forward, backward, intercept, displayEquation });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
        /// <param name="order">optional object order</param>
        /// <param name="period">optional object period</param>
        /// <param name="forward">optional object forward</param>
        /// <param name="backward">optional object backward</param>
        /// <param name="intercept">optional object intercept</param>
        /// <param name="displayEquation">optional object displayEquation</param>
        /// <param name="displayRSquared">optional object displayRSquared</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation, object displayRSquared)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", typeof(NetOffice.OfficeApi.IMsoTrendline), new object[] { type, order, period, forward, backward, intercept, displayEquation, displayRSquared });
        }

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.IMsoTrendline>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.IMsoTrendline>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.IMsoTrendline>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.IMsoTrendline>

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.OfficeApi.IMsoTrendline> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.IMsoTrendline item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsMethod(this, false);
        }

        #endregion

        #pragma warning restore
    }
}
