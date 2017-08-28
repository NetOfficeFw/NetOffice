using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface DiagramNode 
	/// SupportByVersion PowerPoint, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class DiagramNode : COMObject
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

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public DiagramNode(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DiagramNode(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DiagramNode(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DiagramNode(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DiagramNode(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DiagramNode(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DiagramNode() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DiagramNode(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16), ProxyResult]
		public object Application
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.DiagramNodeChildren Children
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.DiagramNodeChildren>(this, "Children", NetOffice.PowerPointApi.DiagramNodeChildren.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape Shape
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Shape>(this, "Shape", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.DiagramNode Root
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.DiagramNode>(this, "Root", NetOffice.PowerPointApi.DiagramNode.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Diagram Diagram
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Diagram>(this, "Diagram", NetOffice.PowerPointApi.Diagram.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoOrgChartLayoutType Layout
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
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape TextShape
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Shape>(this, "TextShape", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = 2</param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoDiagramNodeType NodeType = 1</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.DiagramNode AddNode(object pos, object nodeType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.DiagramNode>(this, "AddNode", NetOffice.PowerPointApi.DiagramNode.LateBindingApiWrapperType, pos, nodeType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.DiagramNode AddNode()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.DiagramNode>(this, "AddNode", NetOffice.PowerPointApi.DiagramNode.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = 2</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.DiagramNode AddNode(object pos)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.DiagramNode>(this, "AddNode", NetOffice.PowerPointApi.DiagramNode.LateBindingApiWrapperType, pos);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="targetNode">NetOffice.PowerPointApi.DiagramNode targetNode</param>
		/// <param name="pos">NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void MoveNode(NetOffice.PowerPointApi.DiagramNode targetNode, NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos)
		{
			 Factory.ExecuteMethod(this, "MoveNode", targetNode, pos);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="targetNode">NetOffice.PowerPointApi.DiagramNode targetNode</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void ReplaceNode(NetOffice.PowerPointApi.DiagramNode targetNode)
		{
			 Factory.ExecuteMethod(this, "ReplaceNode", targetNode);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="targetNode">NetOffice.PowerPointApi.DiagramNode targetNode</param>
		/// <param name="swapChildren">optional bool SwapChildren = true</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void SwapNode(NetOffice.PowerPointApi.DiagramNode targetNode, object swapChildren)
		{
			 Factory.ExecuteMethod(this, "SwapNode", targetNode, swapChildren);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="targetNode">NetOffice.PowerPointApi.DiagramNode targetNode</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void SwapNode(NetOffice.PowerPointApi.DiagramNode targetNode)
		{
			 Factory.ExecuteMethod(this, "SwapNode", targetNode);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="copyChildren">bool copyChildren</param>
		/// <param name="targetNode">NetOffice.PowerPointApi.DiagramNode targetNode</param>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = 2</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.DiagramNode CloneNode(bool copyChildren, NetOffice.PowerPointApi.DiagramNode targetNode, object pos)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.DiagramNode>(this, "CloneNode", NetOffice.PowerPointApi.DiagramNode.LateBindingApiWrapperType, copyChildren, targetNode, pos);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="copyChildren">bool copyChildren</param>
		/// <param name="targetNode">NetOffice.PowerPointApi.DiagramNode targetNode</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.DiagramNode CloneNode(bool copyChildren, NetOffice.PowerPointApi.DiagramNode targetNode)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.DiagramNode>(this, "CloneNode", NetOffice.PowerPointApi.DiagramNode.LateBindingApiWrapperType, copyChildren, targetNode);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="receivingNode">NetOffice.PowerPointApi.DiagramNode receivingNode</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void TransferChildren(NetOffice.PowerPointApi.DiagramNode receivingNode)
		{
			 Factory.ExecuteMethod(this, "TransferChildren", receivingNode);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.DiagramNode NextNode()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.DiagramNode>(this, "NextNode", NetOffice.PowerPointApi.DiagramNode.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.DiagramNode PrevNode()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.DiagramNode>(this, "PrevNode", NetOffice.PowerPointApi.DiagramNode.LateBindingApiWrapperType);
		}

		#endregion

		#pragma warning restore
	}
}
