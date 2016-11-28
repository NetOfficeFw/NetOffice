using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OfficeApi
{
	///<summary>
	/// DispatchInterface DiagramNode 
	/// SupportByVersion Office, 10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class DiagramNode : _IMsoDispObj
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DiagramNode(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.DiagramNodeChildren Children
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Children", paramsArray);
				NetOffice.OfficeApi.DiagramNodeChildren newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.DiagramNodeChildren.LateBindingApiWrapperType) as NetOffice.OfficeApi.DiagramNodeChildren;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape Shape
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shape", paramsArray);
				NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.DiagramNode Root
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Root", paramsArray);
				NetOffice.OfficeApi.DiagramNode newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.DiagramNode.LateBindingApiWrapperType) as NetOffice.OfficeApi.DiagramNode;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.IMsoDiagram Diagram
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Diagram", paramsArray);
				NetOffice.OfficeApi.IMsoDiagram newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoDiagram.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoDiagram;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoOrgChartLayoutType Layout
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Layout", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoOrgChartLayoutType)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Layout", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape TextShape
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TextShape", paramsArray);
				NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = 2</param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoDiagramNodeType NodeType = 1</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.DiagramNode AddNode(object pos, object nodeType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pos, nodeType);
			object returnItem = Invoker.MethodReturn(this, "AddNode", paramsArray);
			NetOffice.OfficeApi.DiagramNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.DiagramNode.LateBindingApiWrapperType) as NetOffice.OfficeApi.DiagramNode;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.DiagramNode AddNode()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AddNode", paramsArray);
			NetOffice.OfficeApi.DiagramNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.DiagramNode.LateBindingApiWrapperType) as NetOffice.OfficeApi.DiagramNode;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = 2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.DiagramNode AddNode(object pos)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pos);
			object returnItem = Invoker.MethodReturn(this, "AddNode", paramsArray);
			NetOffice.OfficeApi.DiagramNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.DiagramNode.LateBindingApiWrapperType) as NetOffice.OfficeApi.DiagramNode;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="targetNode">NetOffice.OfficeApi.DiagramNode TargetNode</param>
		/// <param name="pos">NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void MoveNode(NetOffice.OfficeApi.DiagramNode targetNode, NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(targetNode, pos);
			Invoker.Method(this, "MoveNode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="targetNode">NetOffice.OfficeApi.DiagramNode TargetNode</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void ReplaceNode(NetOffice.OfficeApi.DiagramNode targetNode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(targetNode);
			Invoker.Method(this, "ReplaceNode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="targetNode">NetOffice.OfficeApi.DiagramNode TargetNode</param>
		/// <param name="swapChildren">optional bool SwapChildren = true</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SwapNode(NetOffice.OfficeApi.DiagramNode targetNode, object swapChildren)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(targetNode, swapChildren);
			Invoker.Method(this, "SwapNode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="targetNode">NetOffice.OfficeApi.DiagramNode TargetNode</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SwapNode(NetOffice.OfficeApi.DiagramNode targetNode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(targetNode);
			Invoker.Method(this, "SwapNode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="copyChildren">bool CopyChildren</param>
		/// <param name="targetNode">NetOffice.OfficeApi.DiagramNode TargetNode</param>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = 2</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.DiagramNode CloneNode(bool copyChildren, NetOffice.OfficeApi.DiagramNode targetNode, object pos)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(copyChildren, targetNode, pos);
			object returnItem = Invoker.MethodReturn(this, "CloneNode", paramsArray);
			NetOffice.OfficeApi.DiagramNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.DiagramNode.LateBindingApiWrapperType) as NetOffice.OfficeApi.DiagramNode;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="copyChildren">bool CopyChildren</param>
		/// <param name="targetNode">NetOffice.OfficeApi.DiagramNode TargetNode</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.DiagramNode CloneNode(bool copyChildren, NetOffice.OfficeApi.DiagramNode targetNode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(copyChildren, targetNode);
			object returnItem = Invoker.MethodReturn(this, "CloneNode", paramsArray);
			NetOffice.OfficeApi.DiagramNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.DiagramNode.LateBindingApiWrapperType) as NetOffice.OfficeApi.DiagramNode;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="receivingNode">NetOffice.OfficeApi.DiagramNode ReceivingNode</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void TransferChildren(NetOffice.OfficeApi.DiagramNode receivingNode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(receivingNode);
			Invoker.Method(this, "TransferChildren", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.DiagramNode NextNode()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "NextNode", paramsArray);
			NetOffice.OfficeApi.DiagramNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.DiagramNode.LateBindingApiWrapperType) as NetOffice.OfficeApi.DiagramNode;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.DiagramNode PrevNode()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "PrevNode", paramsArray);
			NetOffice.OfficeApi.DiagramNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.DiagramNode.LateBindingApiWrapperType) as NetOffice.OfficeApi.DiagramNode;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}