using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface DiagramNode 
	/// SupportByVersion Word, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Word", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class DiagramNode : COMObject, NetOffice.WordApi.DiagramNode
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
                    _contractType = typeof(NetOffice.WordApi.DiagramNode);
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
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual Int32 Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.DiagramNodeChildren Children
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.DiagramNodeChildren>(this, "Children", typeof(NetOffice.WordApi.DiagramNodeChildren));
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Shape Shape
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Shape>(this, "Shape", typeof(NetOffice.WordApi.Shape));
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.DiagramNode Root
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.DiagramNode>(this, "Root", typeof(NetOffice.WordApi.DiagramNode));
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Diagram Diagram
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Diagram>(this, "Diagram", typeof(NetOffice.WordApi.Diagram));
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoOrgChartLayoutType Layout
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoOrgChartLayoutType>(this, "Layout");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Layout", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Shape TextShape
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Shape>(this, "TextShape", typeof(NetOffice.WordApi.Shape));
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
		public virtual NetOffice.WordApi.DiagramNode AddNode(object pos, object nodeType)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.DiagramNode>(this, "AddNode", typeof(NetOffice.WordApi.DiagramNode), pos, nodeType);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.DiagramNode AddNode()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.DiagramNode>(this, "AddNode", typeof(NetOffice.WordApi.DiagramNode));
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = 2</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.DiagramNode AddNode(object pos)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.DiagramNode>(this, "AddNode", typeof(NetOffice.WordApi.DiagramNode), pos);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void Delete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="targetNode">NetOffice.WordApi.DiagramNode targetNode</param>
		/// <param name="pos">NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void MoveNode(out NetOffice.WordApi.DiagramNode targetNode, NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,false);
			targetNode = null;
			object[] paramsArray = Invoker.ValidateParamsArray(targetNode, pos);
            InvokerService.InvokeInternal.ExecuteMethod(this, "MoveNode", paramsArray, modifiers);

            if (paramsArray[0] is MarshalByRefObject)
                targetNode = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.DiagramNode>(this, paramsArray[0], typeof(NetOffice.WordApi.DiagramNode));
            else
                targetNode = null;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="targetNode">NetOffice.WordApi.DiagramNode targetNode</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void ReplaceNode(out NetOffice.WordApi.DiagramNode targetNode)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			targetNode = null;
			object[] paramsArray = Invoker.ValidateParamsArray(targetNode);
            InvokerService.InvokeInternal.ExecuteMethod(this, "ReplaceNode", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                targetNode = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.DiagramNode>(this, paramsArray[0], typeof(NetOffice.WordApi.DiagramNode));
            else
                targetNode = null;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="targetNode">NetOffice.WordApi.DiagramNode targetNode</param>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = -1</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void SwapNode(out NetOffice.WordApi.DiagramNode targetNode, object pos)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,false);
			targetNode = null;
			object[] paramsArray = Invoker.ValidateParamsArray(targetNode, pos);
            InvokerService.InvokeInternal.ExecuteMethod(this, "SwapNode", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                targetNode = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.DiagramNode>(this, paramsArray[0], typeof(NetOffice.WordApi.DiagramNode));
            else
                targetNode = null;
        }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="targetNode">NetOffice.WordApi.DiagramNode targetNode</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void SwapNode(out NetOffice.WordApi.DiagramNode targetNode)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			targetNode = null;
			object[] paramsArray = Invoker.ValidateParamsArray(targetNode);
            InvokerService.InvokeInternal.ExecuteMethod(this, "SwapNode", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                targetNode = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.DiagramNode>(this, paramsArray[0], typeof(NetOffice.WordApi.DiagramNode));
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
		public virtual NetOffice.WordApi.DiagramNode CloneNode(bool copyChildren, object targetNode, object pos)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.DiagramNode>(this, "CloneNode", typeof(NetOffice.WordApi.DiagramNode), copyChildren, targetNode, pos);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="copyChildren">bool copyChildren</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.DiagramNode CloneNode(bool copyChildren)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.DiagramNode>(this, "CloneNode", typeof(NetOffice.WordApi.DiagramNode), copyChildren);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="copyChildren">bool copyChildren</param>
		/// <param name="targetNode">optional NetOffice.WordApi.DiagramNode TargetNode = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.DiagramNode CloneNode(bool copyChildren, object targetNode)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.DiagramNode>(this, "CloneNode", typeof(NetOffice.WordApi.DiagramNode), copyChildren, targetNode);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="receivingNode">NetOffice.WordApi.DiagramNode receivingNode</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void TransferChildren(out NetOffice.WordApi.DiagramNode receivingNode)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			receivingNode = null;
			object[] paramsArray = Invoker.ValidateParamsArray(receivingNode);
            InvokerService.InvokeInternal.ExecuteMethod(this, "TransferChildren", paramsArray, modifiers);
			receivingNode = (NetOffice.WordApi.DiagramNode)paramsArray[0];
            if (paramsArray[0] is MarshalByRefObject)
                receivingNode = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.DiagramNode>(this, paramsArray[0], typeof(NetOffice.WordApi.DiagramNode));
            else
                receivingNode = null;
        }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.DiagramNode NextNode()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.DiagramNode>(this, "NextNode", typeof(NetOffice.WordApi.DiagramNode));
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.DiagramNode PrevNode()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.DiagramNode>(this, "PrevNode", typeof(NetOffice.WordApi.DiagramNode));
		}

		#endregion

		#pragma warning restore
	}
}


