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
    /// DispatchInterface DiagramNodeChildren 
    /// SupportByVersion Office, 10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
    [Duplicate("NetOffice.ExcelApi.DiagramNodeChildren")]
    public class DiagramNodeChildren : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.DiagramNodeChildren
    {
        #pragma warning disable

        #region Type Information

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
                    _type = typeof(DiagramNodeChildren);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public DiagramNodeChildren() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16), ProxyResult]
        public object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public Int32 Count
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.DiagramNode FirstChild
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DiagramNode>(this, "FirstChild", typeof(NetOffice.OfficeApi.DiagramNode));
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.DiagramNode LastChild
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DiagramNode>(this, "LastChild", typeof(NetOffice.OfficeApi.DiagramNode));
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public NetOffice.OfficeApi.DiagramNode this[object index]
        {
            get
            {
                return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.DiagramNode>(this, "Item", typeof(NetOffice.OfficeApi.DiagramNode), index);
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object Index = -1</param>
        /// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoDiagramNodeType NodeType = 1</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.DiagramNode AddNode(object index, object nodeType)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.DiagramNode>(this, "AddNode", typeof(NetOffice.OfficeApi.DiagramNode), index, nodeType);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.DiagramNode AddNode()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.DiagramNode>(this, "AddNode", typeof(NetOffice.OfficeApi.DiagramNode));
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object Index = -1</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.DiagramNode AddNode(object index)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.DiagramNode>(this, "AddNode", typeof(NetOffice.OfficeApi.DiagramNode), index);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public void SelectAll()
        {
            Factory.ExecuteMethod(this, "SelectAll");
        }

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.DiagramNode>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.DiagramNode>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.DiagramNode>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.DiagramNode>

        /// <summary>
        /// SupportByVersion Office, 10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.DiagramNode> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.DiagramNode item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
        }

        #endregion

        #pragma warning restore
    }
}
