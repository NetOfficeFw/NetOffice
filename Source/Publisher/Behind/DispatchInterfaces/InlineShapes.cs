using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PublisherApi.Behind
{
    /// <summary>
    /// InlineShapes
    /// </summary>
    [SyntaxBypass]
    public class InlineShapes_ : COMObject
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public InlineShapes_() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.PublisherApi.ShapeRange get_Range(object index)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ShapeRange>(this, "Range", typeof(NetOffice.PublisherApi.ShapeRange), index);
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Alias for get_Range
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Publisher", 14, 15, 16), Redirect("get_Range")]
        public NetOffice.PublisherApi.ShapeRange Range(object index)
        {
            return get_Range(index);
        }

        #endregion
    }

    /// <summary>
    /// DispatchInterface InlineShapes 
    /// SupportByVersion Publisher, 14,15,16
    /// </summary>
     public class InlineShapes : InlineShapes_, NetOffice.PublisherApi.InlineShapes
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
                    _contractType = typeof(NetOffice.PublisherApi.InlineShapes);
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
                    _type = typeof(InlineShapes);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public InlineShapes() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public NetOffice.PublisherApi.Application Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", typeof(NetOffice.PublisherApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16), ProxyResult]
        public object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public Int32 Count
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public NetOffice.PublisherApi.ShapeRange Range
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ShapeRange>(this, "Range", typeof(NetOffice.PublisherApi.ShapeRange));
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Publisher 14, 15, 16
        /// </summary>
        /// <param name="var">object var</param>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public NetOffice.PublisherApi.Shape this[object var]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "Item", typeof(NetOffice.PublisherApi.Shape), var);
            }
        }

        #endregion

        #region IEnumerableProvider<NetOffice.PublisherApi.Shape>

        ICOMObject IEnumerableProvider<NetOffice.PublisherApi.Shape>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.PublisherApi.Shape>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.PublisherApi.Shape>

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public IEnumerator<NetOffice.PublisherApi.Shape> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.PublisherApi.Shape item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
        }

        #endregion

        #pragma warning restore
    }
}
