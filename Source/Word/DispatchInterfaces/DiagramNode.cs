using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface DiagramNode 
	/// SupportByVersion Word, 10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class DiagramNode : COMObject
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
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
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
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNodeChildren Children
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Children", paramsArray);
				NetOffice.WordApi.DiagramNodeChildren newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.DiagramNodeChildren.LateBindingApiWrapperType) as NetOffice.WordApi.DiagramNodeChildren;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape Shape
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shape", paramsArray);
				NetOffice.WordApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Shape.LateBindingApiWrapperType) as NetOffice.WordApi.Shape;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNode Root
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Root", paramsArray);
				NetOffice.WordApi.DiagramNode newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.DiagramNode.LateBindingApiWrapperType) as NetOffice.WordApi.DiagramNode;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Diagram Diagram
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Diagram", paramsArray);
				NetOffice.WordApi.Diagram newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Diagram.LateBindingApiWrapperType) as NetOffice.WordApi.Diagram;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
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
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape TextShape
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TextShape", paramsArray);
				NetOffice.WordApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Shape.LateBindingApiWrapperType) as NetOffice.WordApi.Shape;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = 2</param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoDiagramNodeType NodeType = 1</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNode AddNode(object pos, object nodeType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pos, nodeType);
			object returnItem = Invoker.MethodReturn(this, "AddNode", paramsArray);
			NetOffice.WordApi.DiagramNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.DiagramNode.LateBindingApiWrapperType) as NetOffice.WordApi.DiagramNode;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNode AddNode()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AddNode", paramsArray);
			NetOffice.WordApi.DiagramNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.DiagramNode.LateBindingApiWrapperType) as NetOffice.WordApi.DiagramNode;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = 2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNode AddNode(object pos)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pos);
			object returnItem = Invoker.MethodReturn(this, "AddNode", paramsArray);
			NetOffice.WordApi.DiagramNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.DiagramNode.LateBindingApiWrapperType) as NetOffice.WordApi.DiagramNode;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="targetNode">NetOffice.WordApi.DiagramNode TargetNode</param>
		/// <param name="pos">NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void MoveNode(out NetOffice.WordApi.DiagramNode targetNode, NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,false);
			targetNode = null;
			object[] paramsArray = Invoker.ValidateParamsArray(targetNode, pos);
			Invoker.Method(this, "MoveNode", paramsArray, modifiers);
			targetNode = (NetOffice.WordApi.DiagramNode)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="targetNode">NetOffice.WordApi.DiagramNode TargetNode</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void ReplaceNode(out NetOffice.WordApi.DiagramNode targetNode)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			targetNode = null;
			object[] paramsArray = Invoker.ValidateParamsArray(targetNode);
			Invoker.Method(this, "ReplaceNode", paramsArray, modifiers);
			targetNode = (NetOffice.WordApi.DiagramNode)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="targetNode">NetOffice.WordApi.DiagramNode TargetNode</param>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = -1</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SwapNode(out NetOffice.WordApi.DiagramNode targetNode, object pos)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,false);
			targetNode = null;
			object[] paramsArray = Invoker.ValidateParamsArray(targetNode, pos);
			Invoker.Method(this, "SwapNode", paramsArray, modifiers);
			targetNode = (NetOffice.WordApi.DiagramNode)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="targetNode">NetOffice.WordApi.DiagramNode TargetNode</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void SwapNode(out NetOffice.WordApi.DiagramNode targetNode)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			targetNode = null;
			object[] paramsArray = Invoker.ValidateParamsArray(targetNode);
			Invoker.Method(this, "SwapNode", paramsArray, modifiers);
			targetNode = (NetOffice.WordApi.DiagramNode)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="copyChildren">bool copyChildren</param>
		/// <param name="targetNode">optional NetOffice.WordApi.DiagramNode TargetNode = 0</param>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = 2</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNode CloneNode(bool copyChildren, object targetNode, object pos)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(copyChildren, targetNode, pos);
			object returnItem = Invoker.MethodReturn(this, "CloneNode", paramsArray);
			NetOffice.WordApi.DiagramNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.DiagramNode.LateBindingApiWrapperType) as NetOffice.WordApi.DiagramNode;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="copyChildren">bool copyChildren</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNode CloneNode(bool copyChildren)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(copyChildren);
			object returnItem = Invoker.MethodReturn(this, "CloneNode", paramsArray);
			NetOffice.WordApi.DiagramNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.DiagramNode.LateBindingApiWrapperType) as NetOffice.WordApi.DiagramNode;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="copyChildren">bool copyChildren</param>
		/// <param name="targetNode">optional NetOffice.WordApi.DiagramNode TargetNode = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNode CloneNode(bool copyChildren, object targetNode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(copyChildren, targetNode);
			object returnItem = Invoker.MethodReturn(this, "CloneNode", paramsArray);
			NetOffice.WordApi.DiagramNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.DiagramNode.LateBindingApiWrapperType) as NetOffice.WordApi.DiagramNode;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="receivingNode">NetOffice.WordApi.DiagramNode ReceivingNode</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void TransferChildren(out NetOffice.WordApi.DiagramNode receivingNode)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			receivingNode = null;
			object[] paramsArray = Invoker.ValidateParamsArray(receivingNode);
			Invoker.Method(this, "TransferChildren", paramsArray, modifiers);
			receivingNode = (NetOffice.WordApi.DiagramNode)paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNode NextNode()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "NextNode", paramsArray);
			NetOffice.WordApi.DiagramNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.DiagramNode.LateBindingApiWrapperType) as NetOffice.WordApi.DiagramNode;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.DiagramNode PrevNode()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "PrevNode", paramsArray);
			NetOffice.WordApi.DiagramNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.DiagramNode.LateBindingApiWrapperType) as NetOffice.WordApi.DiagramNode;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}