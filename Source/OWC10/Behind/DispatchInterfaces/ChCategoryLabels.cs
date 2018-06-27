using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OWC10Api.Behind
{
    /// <summary>
    /// ChCategoryLabels
    /// </summary>
    [SyntaxBypass]
    public class ChCategoryLabels_ : COMObject
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public ChCategoryLabels_() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="level">optional Int32 level</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 get_ItemCount(object level)
        {
            return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ItemCount", level);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_ItemCount
        /// </summary>
        /// <param name="level">optional Int32 level</param>
        [SupportByVersion("OWC10", 1), Redirect("get_ItemCount")]
        public virtual Int32 ItemCount(object level)
        {
            return get_ItemCount(level);
        }

        #endregion
    }

    /// <summary>
    /// DispatchInterface ChCategoryLabels 
    /// SupportByVersion OWC10, 1
    /// </summary>
    public class ChCategoryLabels : ChCategoryLabels_, NetOffice.OWC10Api.ChCategoryLabels
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
                    _contractType = typeof(NetOffice.OWC10Api.ChCategoryLabels);
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
                    _type = typeof(ChCategoryLabels);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public ChCategoryLabels() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual Int32 LevelCount
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "LevelCount");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual Int32 ItemCount
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ItemCount");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.ChAxis Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChAxis>(this, "Parent", typeof(NetOffice.OWC10Api.ChAxis));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// Custom Indexer
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("OWC10", 1)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
        public virtual NetOffice.OWC10Api.ChCategoryLabel this[object index]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChCategoryLabel>(this, "Item", typeof(NetOffice.OWC10Api.ChCategoryLabel), index);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="index">object index</param>
        /// <param name="level">optional Int32 level</param>
        [SupportByVersion("OWC10", 1)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public virtual NetOffice.OWC10Api.ChCategoryLabel this[object index, object level]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChCategoryLabel>(this, "Item", typeof(NetOffice.OWC10Api.ChCategoryLabel), index, level);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        public virtual NetOffice.OWC10Api.PivotResultGroupAxis PivotAxis
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.PivotResultGroupAxis>(this, "PivotAxis");
            }
        }

        #endregion

        #region IEnumerableProvider<NetOffice.OWC10Api.ChCategoryLabel>

        ICOMObject IEnumerableProvider<NetOffice.OWC10Api.ChCategoryLabel>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.OWC10Api.ChCategoryLabel>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.ChCategoryLabel>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual IEnumerator<NetOffice.OWC10Api.ChCategoryLabel> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OWC10Api.ChCategoryLabel item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
        }

        #endregion

        #pragma warning restore
    }
}
