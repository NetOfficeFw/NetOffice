using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
	/// <summary>
	/// DispatchInterface DiagramNode 
	/// SupportByVersion Excel, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
    [Duplicate("NetOffice.OfficeApi.DiagramNode")]
    public class DiagramNode : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.ExcelApi.DiagramNode
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.DiagramNodeChildren Children
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.DiagramNodeChildren>(this, "Children", typeof(NetOffice.ExcelApi.DiagramNodeChildren));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Shape Shape
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Shape>(this, "Shape", typeof(NetOffice.ExcelApi.Shape));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.DiagramNode Root
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.DiagramNode>(this, "Root", typeof(NetOffice.ExcelApi.DiagramNode));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.IMsoDiagram Diagram
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoDiagram>(this, "Diagram", typeof(NetOffice.OfficeApi.IMsoDiagram));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Shape TextShape
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Shape>(this, "TextShape", typeof(NetOffice.ExcelApi.Shape));
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos = 2</param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoDiagramNodeType nodeType = 1</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.DiagramNode AddNode(object pos, object nodeType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.DiagramNode>(this, "AddNode", typeof(NetOffice.ExcelApi.DiagramNode), pos, nodeType);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.DiagramNode AddNode()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.DiagramNode>(this, "AddNode", typeof(NetOffice.ExcelApi.DiagramNode));
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos = 2</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.DiagramNode AddNode(object pos)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.DiagramNode>(this, "AddNode", typeof(NetOffice.ExcelApi.DiagramNode), pos);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pTargetNode">NetOffice.ExcelApi.DiagramNode pTargetNode</param>
		/// <param name="pos">NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual void MoveNode(NetOffice.ExcelApi.DiagramNode pTargetNode, NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos)
		{
			 Factory.ExecuteMethod(this, "MoveNode", pTargetNode, pos);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pTargetNode">NetOffice.ExcelApi.DiagramNode pTargetNode</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual void ReplaceNode(NetOffice.ExcelApi.DiagramNode pTargetNode)
		{
			 Factory.ExecuteMethod(this, "ReplaceNode", pTargetNode);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pTargetNode">NetOffice.ExcelApi.DiagramNode pTargetNode</param>
		/// <param name="swapChildren">optional bool swapChildren = true</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual void SwapNode(NetOffice.ExcelApi.DiagramNode pTargetNode, object swapChildren)
		{
			 Factory.ExecuteMethod(this, "SwapNode", pTargetNode, swapChildren);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pTargetNode">NetOffice.ExcelApi.DiagramNode pTargetNode</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual void SwapNode(NetOffice.ExcelApi.DiagramNode pTargetNode)
		{
			 Factory.ExecuteMethod(this, "SwapNode", pTargetNode);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="copyChildren">bool copyChildren</param>
		/// <param name="pTargetNode">NetOffice.ExcelApi.DiagramNode pTargetNode</param>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos = 2</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.DiagramNode CloneNode(bool copyChildren, NetOffice.ExcelApi.DiagramNode pTargetNode, object pos)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.DiagramNode>(this, "CloneNode", typeof(NetOffice.ExcelApi.DiagramNode), copyChildren, pTargetNode, pos);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="copyChildren">bool copyChildren</param>
		/// <param name="pTargetNode">NetOffice.ExcelApi.DiagramNode pTargetNode</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.DiagramNode CloneNode(bool copyChildren, NetOffice.ExcelApi.DiagramNode pTargetNode)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.DiagramNode>(this, "CloneNode", typeof(NetOffice.ExcelApi.DiagramNode), copyChildren, pTargetNode);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pReceivingNode">NetOffice.ExcelApi.DiagramNode pReceivingNode</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual void TransferChildren(NetOffice.ExcelApi.DiagramNode pReceivingNode)
		{
			 Factory.ExecuteMethod(this, "TransferChildren", pReceivingNode);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.DiagramNode NextNode()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.DiagramNode>(this, "NextNode", typeof(NetOffice.ExcelApi.DiagramNode));
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.DiagramNode PrevNode()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.DiagramNode>(this, "PrevNode", typeof(NetOffice.ExcelApi.DiagramNode));
		}

		#endregion

		#pragma warning restore
	}
}


