using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface DiagramNode 
    /// SupportByVersion Office, 10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    [Duplicate("NetOffice.ExcelApi.DiagramNode")]
    public class DiagramNode : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.DiagramNode
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
                    _type = typeof(DiagramNode);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public DiagramNode() : base()
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
        public virtual object Parent
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
        public virtual NetOffice.OfficeApi.DiagramNodeChildren Children
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DiagramNodeChildren>(this, "Children", typeof(NetOffice.OfficeApi.DiagramNodeChildren));
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Shape Shape
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Shape>(this, "Shape", typeof(NetOffice.OfficeApi.Shape));
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.DiagramNode Root
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DiagramNode>(this, "Root", typeof(NetOffice.OfficeApi.DiagramNode));
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoDiagram Diagram
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoDiagram>(this, "Diagram", typeof(NetOffice.OfficeApi.IMsoDiagram));
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoOrgChartLayoutType Layout
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoOrgChartLayoutType>(this, "Layout");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "Layout", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Shape TextShape
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Shape>(this, "TextShape", typeof(NetOffice.OfficeApi.Shape));
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = 2</param>
        /// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoDiagramNodeType NodeType = 1</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.DiagramNode AddNode(object pos, object nodeType)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.DiagramNode>(this, "AddNode", typeof(NetOffice.OfficeApi.DiagramNode), pos, nodeType);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.DiagramNode AddNode()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.DiagramNode>(this, "AddNode", typeof(NetOffice.OfficeApi.DiagramNode));
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = 2</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.DiagramNode AddNode(object pos)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.DiagramNode>(this, "AddNode", typeof(NetOffice.OfficeApi.DiagramNode), pos);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void Delete()
        {
            Factory.ExecuteMethod(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="targetNode">NetOffice.OfficeApi.DiagramNode targetNode</param>
        /// <param name="pos">NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void MoveNode(NetOffice.OfficeApi.DiagramNode targetNode, NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos)
        {
            Factory.ExecuteMethod(this, "MoveNode", targetNode, pos);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="targetNode">NetOffice.OfficeApi.DiagramNode targetNode</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void ReplaceNode(NetOffice.OfficeApi.DiagramNode targetNode)
        {
            Factory.ExecuteMethod(this, "ReplaceNode", targetNode);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="targetNode">NetOffice.OfficeApi.DiagramNode targetNode</param>
        /// <param name="swapChildren">optional bool SwapChildren = true</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SwapNode(NetOffice.OfficeApi.DiagramNode targetNode, object swapChildren)
        {
            Factory.ExecuteMethod(this, "SwapNode", targetNode, swapChildren);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="targetNode">NetOffice.OfficeApi.DiagramNode targetNode</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SwapNode(NetOffice.OfficeApi.DiagramNode targetNode)
        {
            Factory.ExecuteMethod(this, "SwapNode", targetNode);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="copyChildren">bool copyChildren</param>
        /// <param name="targetNode">NetOffice.OfficeApi.DiagramNode targetNode</param>
        /// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = 2</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.DiagramNode CloneNode(bool copyChildren, NetOffice.OfficeApi.DiagramNode targetNode, object pos)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.DiagramNode>(this, "CloneNode", typeof(NetOffice.OfficeApi.DiagramNode), copyChildren, targetNode, pos);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="copyChildren">bool copyChildren</param>
        /// <param name="targetNode">NetOffice.OfficeApi.DiagramNode targetNode</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.DiagramNode CloneNode(bool copyChildren, NetOffice.OfficeApi.DiagramNode targetNode)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.DiagramNode>(this, "CloneNode", typeof(NetOffice.OfficeApi.DiagramNode), copyChildren, targetNode);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="receivingNode">NetOffice.OfficeApi.DiagramNode receivingNode</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void TransferChildren(NetOffice.OfficeApi.DiagramNode receivingNode)
        {
            Factory.ExecuteMethod(this, "TransferChildren", receivingNode);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.DiagramNode NextNode()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.DiagramNode>(this, "NextNode", typeof(NetOffice.OfficeApi.DiagramNode));
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.DiagramNode PrevNode()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.DiagramNode>(this, "PrevNode", typeof(NetOffice.OfficeApi.DiagramNode));
        }

        #endregion

        #pragma warning restore
    }
}
