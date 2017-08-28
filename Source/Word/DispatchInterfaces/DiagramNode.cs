using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface DiagramNode 
	/// SupportByVersion Word, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Word", 10,11,12,14,15,16)]
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
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNodeChildren Children
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.DiagramNodeChildren>(this, "Children", NetOffice.WordApi.DiagramNodeChildren.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape Shape
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Shape>(this, "Shape", NetOffice.WordApi.Shape.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNode Root
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.DiagramNode>(this, "Root", NetOffice.WordApi.DiagramNode.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Diagram Diagram
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Diagram>(this, "Diagram", NetOffice.WordApi.Diagram.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
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
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape TextShape
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Shape>(this, "TextShape", NetOffice.WordApi.Shape.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = 2</param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoDiagramNodeType NodeType = 1</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNode AddNode(object pos, object nodeType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.DiagramNode>(this, "AddNode", NetOffice.WordApi.DiagramNode.LateBindingApiWrapperType, pos, nodeType);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNode AddNode()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.DiagramNode>(this, "AddNode", NetOffice.WordApi.DiagramNode.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = 2</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNode AddNode(object pos)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.DiagramNode>(this, "AddNode", NetOffice.WordApi.DiagramNode.LateBindingApiWrapperType, pos);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="targetNode">NetOffice.WordApi.DiagramNode targetNode</param>
		/// <param name="pos">NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void MoveNode(out NetOffice.WordApi.DiagramNode targetNode, NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,false);
			targetNode = null;
			object[] paramsArray = Invoker.ValidateParamsArray(targetNode, pos);
			Invoker.Method(this, "MoveNode", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                targetNode = new NetOffice.WordApi.DiagramNode(this, paramsArray[0]);
            else
                targetNode = null;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="targetNode">NetOffice.WordApi.DiagramNode targetNode</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void ReplaceNode(out NetOffice.WordApi.DiagramNode targetNode)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			targetNode = null;
			object[] paramsArray = Invoker.ValidateParamsArray(targetNode);
			Invoker.Method(this, "ReplaceNode", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                targetNode = new NetOffice.WordApi.DiagramNode(this, paramsArray[0]);
            else
                targetNode = null;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="targetNode">NetOffice.WordApi.DiagramNode targetNode</param>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = -1</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SwapNode(out NetOffice.WordApi.DiagramNode targetNode, object pos)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,false);
			targetNode = null;
			object[] paramsArray = Invoker.ValidateParamsArray(targetNode, pos);
			Invoker.Method(this, "SwapNode", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                targetNode = new NetOffice.WordApi.DiagramNode(this, paramsArray[0]);
            else
                targetNode = null;
        }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="targetNode">NetOffice.WordApi.DiagramNode targetNode</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SwapNode(out NetOffice.WordApi.DiagramNode targetNode)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			targetNode = null;
			object[] paramsArray = Invoker.ValidateParamsArray(targetNode);
			Invoker.Method(this, "SwapNode", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                targetNode = new NetOffice.WordApi.DiagramNode(this, paramsArray[0]);
            else
                targetNode = null;
        }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="copyChildren">bool copyChildren</param>
		/// <param name="targetNode">optional NetOffice.WordApi.DiagramNode TargetNode = 0</param>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = 2</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNode CloneNode(bool copyChildren, object targetNode, object pos)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.DiagramNode>(this, "CloneNode", NetOffice.WordApi.DiagramNode.LateBindingApiWrapperType, copyChildren, targetNode, pos);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="copyChildren">bool copyChildren</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNode CloneNode(bool copyChildren)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.DiagramNode>(this, "CloneNode", NetOffice.WordApi.DiagramNode.LateBindingApiWrapperType, copyChildren);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="copyChildren">bool copyChildren</param>
		/// <param name="targetNode">optional NetOffice.WordApi.DiagramNode TargetNode = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNode CloneNode(bool copyChildren, object targetNode)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.DiagramNode>(this, "CloneNode", NetOffice.WordApi.DiagramNode.LateBindingApiWrapperType, copyChildren, targetNode);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="receivingNode">NetOffice.WordApi.DiagramNode receivingNode</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void TransferChildren(out NetOffice.WordApi.DiagramNode receivingNode)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			receivingNode = null;
			object[] paramsArray = Invoker.ValidateParamsArray(receivingNode);
			Invoker.Method(this, "TransferChildren", paramsArray, modifiers);
			receivingNode = (NetOffice.WordApi.DiagramNode)paramsArray[0];
            if (paramsArray[0] is MarshalByRefObject)
                receivingNode = new NetOffice.WordApi.DiagramNode(this, paramsArray[0]);
            else
                receivingNode = null;
        }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNode NextNode()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.DiagramNode>(this, "NextNode", NetOffice.WordApi.DiagramNode.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNode PrevNode()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.DiagramNode>(this, "PrevNode", NetOffice.WordApi.DiagramNode.LateBindingApiWrapperType);
		}

		#endregion

		#pragma warning restore
	}
}
