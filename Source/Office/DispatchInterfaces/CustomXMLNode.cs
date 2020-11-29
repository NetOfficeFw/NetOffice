using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface CustomXMLNode 
	/// SupportByVersion Office, 12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode"/> </remarks>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class CustomXMLNode : _IMsoDispObj
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
                    _type = typeof(CustomXMLNode);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public CustomXMLNode(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public CustomXMLNode(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomXMLNode(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomXMLNode(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomXMLNode(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomXMLNode(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomXMLNode() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomXMLNode(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.Parent"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.Attributes"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLNodes Attributes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLNodes>(this, "Attributes", NetOffice.OfficeApi.CustomXMLNodes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.BaseName"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public string BaseName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BaseName");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.ChildNodes"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLNodes ChildNodes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLNodes>(this, "ChildNodes", NetOffice.OfficeApi.CustomXMLNodes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.FirstChild"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLNode FirstChild
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLNode>(this, "FirstChild", NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.LastChild"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLNode LastChild
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLNode>(this, "LastChild", NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.NamespaceURI"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public string NamespaceURI
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NamespaceURI");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.NextSibling"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLNode NextSibling
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLNode>(this, "NextSibling", NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.NodeType"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType>(this, "NodeType");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.NodeValue"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public string NodeValue
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NodeValue");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NodeValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.OwnerDocument"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16), ProxyResult]
		public object OwnerDocument
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "OwnerDocument");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.OwnerPart"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLPart OwnerPart
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLPart>(this, "OwnerPart", NetOffice.OfficeApi.CustomXMLPart.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.PreviousSibling"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLNode PreviousSibling
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLNode>(this, "PreviousSibling", NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.ParentNode"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLNode ParentNode
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLNode>(this, "ParentNode", NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.Text"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public string Text
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.XPath"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public string XPath
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "XPath");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.XML"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public string XML
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "XML");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.AppendChildNode"/> </remarks>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
		/// <param name="nodeValue">optional string NodeValue = </param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void AppendChildNode(object name, object namespaceURI, object nodeType, object nodeValue)
		{
			 Factory.ExecuteMethod(this, "AppendChildNode", name, namespaceURI, nodeType, nodeValue);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.AppendChildNode"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void AppendChildNode()
		{
			 Factory.ExecuteMethod(this, "AppendChildNode");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.AppendChildNode"/> </remarks>
		/// <param name="name">optional string Name = </param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void AppendChildNode(object name)
		{
			 Factory.ExecuteMethod(this, "AppendChildNode", name);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.AppendChildNode"/> </remarks>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void AppendChildNode(object name, object namespaceURI)
		{
			 Factory.ExecuteMethod(this, "AppendChildNode", name, namespaceURI);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.AppendChildNode"/> </remarks>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void AppendChildNode(object name, object namespaceURI, object nodeType)
		{
			 Factory.ExecuteMethod(this, "AppendChildNode", name, namespaceURI, nodeType);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.AppendChildSubtree"/> </remarks>
		/// <param name="xML">string xML</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void AppendChildSubtree(string xML)
		{
			 Factory.ExecuteMethod(this, "AppendChildSubtree", xML);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.Delete"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.HasChildNodes"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool HasChildNodes()
		{
			return Factory.ExecuteBoolMethodGet(this, "HasChildNodes");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.InsertNodeBefore"/> </remarks>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
		/// <param name="nodeValue">optional string NodeValue = </param>
		/// <param name="nextSibling">optional NetOffice.OfficeApi.CustomXMLNode NextSibling = 0</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void InsertNodeBefore(object name, object namespaceURI, object nodeType, object nodeValue, object nextSibling)
		{
			 Factory.ExecuteMethod(this, "InsertNodeBefore", new object[]{ name, namespaceURI, nodeType, nodeValue, nextSibling });
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.InsertNodeBefore"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void InsertNodeBefore()
		{
			 Factory.ExecuteMethod(this, "InsertNodeBefore");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.InsertNodeBefore"/> </remarks>
		/// <param name="name">optional string Name = </param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void InsertNodeBefore(object name)
		{
			 Factory.ExecuteMethod(this, "InsertNodeBefore", name);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.InsertNodeBefore"/> </remarks>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void InsertNodeBefore(object name, object namespaceURI)
		{
			 Factory.ExecuteMethod(this, "InsertNodeBefore", name, namespaceURI);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.InsertNodeBefore"/> </remarks>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void InsertNodeBefore(object name, object namespaceURI, object nodeType)
		{
			 Factory.ExecuteMethod(this, "InsertNodeBefore", name, namespaceURI, nodeType);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.InsertNodeBefore"/> </remarks>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
		/// <param name="nodeValue">optional string NodeValue = </param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void InsertNodeBefore(object name, object namespaceURI, object nodeType, object nodeValue)
		{
			 Factory.ExecuteMethod(this, "InsertNodeBefore", name, namespaceURI, nodeType, nodeValue);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.InsertSubtreeBefore"/> </remarks>
		/// <param name="xML">string xML</param>
		/// <param name="nextSibling">optional NetOffice.OfficeApi.CustomXMLNode NextSibling = 0</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void InsertSubtreeBefore(string xML, object nextSibling)
		{
			 Factory.ExecuteMethod(this, "InsertSubtreeBefore", xML, nextSibling);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.InsertSubtreeBefore"/> </remarks>
		/// <param name="xML">string xML</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void InsertSubtreeBefore(string xML)
		{
			 Factory.ExecuteMethod(this, "InsertSubtreeBefore", xML);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.RemoveChild"/> </remarks>
		/// <param name="child">NetOffice.OfficeApi.CustomXMLNode child</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void RemoveChild(NetOffice.OfficeApi.CustomXMLNode child)
		{
			 Factory.ExecuteMethod(this, "RemoveChild", child);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.ReplaceChildNode"/> </remarks>
		/// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
		/// <param name="nodeValue">optional string NodeValue = </param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode, object name, object namespaceURI, object nodeType, object nodeValue)
		{
			 Factory.ExecuteMethod(this, "ReplaceChildNode", new object[]{ oldNode, name, namespaceURI, nodeType, nodeValue });
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.ReplaceChildNode"/> </remarks>
		/// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode)
		{
			 Factory.ExecuteMethod(this, "ReplaceChildNode", oldNode);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.ReplaceChildNode"/> </remarks>
		/// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
		/// <param name="name">optional string Name = </param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode, object name)
		{
			 Factory.ExecuteMethod(this, "ReplaceChildNode", oldNode, name);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.ReplaceChildNode"/> </remarks>
		/// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode, object name, object namespaceURI)
		{
			 Factory.ExecuteMethod(this, "ReplaceChildNode", oldNode, name, namespaceURI);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.ReplaceChildNode"/> </remarks>
		/// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode, object name, object namespaceURI, object nodeType)
		{
			 Factory.ExecuteMethod(this, "ReplaceChildNode", oldNode, name, namespaceURI, nodeType);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.ReplaceChildSubtree"/> </remarks>
		/// <param name="xML">string xML</param>
		/// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void ReplaceChildSubtree(string xML, NetOffice.OfficeApi.CustomXMLNode oldNode)
		{
			 Factory.ExecuteMethod(this, "ReplaceChildSubtree", xML, oldNode);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.SelectNodes"/> </remarks>
		/// <param name="xPath">string xPath</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLNodes SelectNodes(string xPath)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLNodes>(this, "SelectNodes", NetOffice.OfficeApi.CustomXMLNodes.LateBindingApiWrapperType, xPath);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLNode.SelectSingleNode"/> </remarks>
		/// <param name="xPath">string xPath</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLNode SelectSingleNode(string xPath)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLNode>(this, "SelectSingleNode", NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType, xPath);
		}

		#endregion

		#pragma warning restore
	}
}
