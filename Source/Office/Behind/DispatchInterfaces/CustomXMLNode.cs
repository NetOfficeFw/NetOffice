using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface CustomXMLNode 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865242.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class CustomXMLNode : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.CustomXMLNode
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
                    _contractType = typeof(NetOffice.OfficeApi.CustomXMLNode);
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
                    _type = typeof(CustomXMLNode);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public CustomXMLNode() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864640.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861370.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLNodes Attributes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLNodes>(this, "Attributes", typeof(NetOffice.OfficeApi.CustomXMLNodes));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862737.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string BaseName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BaseName");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862357.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLNodes ChildNodes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLNodes>(this, "ChildNodes", typeof(NetOffice.OfficeApi.CustomXMLNodes));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863022.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLNode FirstChild
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLNode>(this, "FirstChild", typeof(NetOffice.OfficeApi.CustomXMLNode));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864028.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLNode LastChild
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLNode>(this, "LastChild", typeof(NetOffice.OfficeApi.CustomXMLNode));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861516.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string NamespaceURI
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "NamespaceURI");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862522.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLNode NextSibling
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLNode>(this, "NextSibling", typeof(NetOffice.OfficeApi.CustomXMLNode));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865216.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType>(this, "NodeType");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862159.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string NodeValue
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "NodeValue");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NodeValue", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862788.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        public virtual object OwnerDocument
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "OwnerDocument");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864973.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLPart OwnerPart
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLPart>(this, "OwnerPart", typeof(NetOffice.OfficeApi.CustomXMLPart));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861743.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLNode PreviousSibling
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLNode>(this, "PreviousSibling", typeof(NetOffice.OfficeApi.CustomXMLNode));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865519.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLNode ParentNode
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLNode>(this, "ParentNode", typeof(NetOffice.OfficeApi.CustomXMLNode));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863358.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string Text
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860871.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string XPath
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "XPath");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860882.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string XML
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "XML");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861364.aspx </remarks>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        /// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
        /// <param name="nodeValue">optional string NodeValue = </param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void AppendChildNode(object name, object namespaceURI, object nodeType, object nodeValue)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AppendChildNode", name, namespaceURI, nodeType, nodeValue);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861364.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void AppendChildNode()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AppendChildNode");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861364.aspx </remarks>
        /// <param name="name">optional string Name = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void AppendChildNode(object name)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AppendChildNode", name);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861364.aspx </remarks>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void AppendChildNode(object name, object namespaceURI)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AppendChildNode", name, namespaceURI);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861364.aspx </remarks>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        /// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void AppendChildNode(object name, object namespaceURI, object nodeType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AppendChildNode", name, namespaceURI, nodeType);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862169.aspx </remarks>
        /// <param name="xML">string xML</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void AppendChildSubtree(string xML)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AppendChildSubtree", xML);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864986.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Delete()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863303.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool HasChildNodes()
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HasChildNodes");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863860.aspx </remarks>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        /// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
        /// <param name="nodeValue">optional string NodeValue = </param>
        /// <param name="nextSibling">optional NetOffice.OfficeApi.CustomXMLNode NextSibling = 0</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void InsertNodeBefore(object name, object namespaceURI, object nodeType, object nodeValue, object nextSibling)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertNodeBefore", new object[] { name, namespaceURI, nodeType, nodeValue, nextSibling });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863860.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void InsertNodeBefore()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertNodeBefore");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863860.aspx </remarks>
        /// <param name="name">optional string Name = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void InsertNodeBefore(object name)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertNodeBefore", name);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863860.aspx </remarks>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void InsertNodeBefore(object name, object namespaceURI)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertNodeBefore", name, namespaceURI);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863860.aspx </remarks>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        /// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void InsertNodeBefore(object name, object namespaceURI, object nodeType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertNodeBefore", name, namespaceURI, nodeType);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863860.aspx </remarks>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        /// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
        /// <param name="nodeValue">optional string NodeValue = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void InsertNodeBefore(object name, object namespaceURI, object nodeType, object nodeValue)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertNodeBefore", name, namespaceURI, nodeType, nodeValue);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861904.aspx </remarks>
        /// <param name="xML">string xML</param>
        /// <param name="nextSibling">optional NetOffice.OfficeApi.CustomXMLNode NextSibling = 0</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void InsertSubtreeBefore(string xML, object nextSibling)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertSubtreeBefore", xML, nextSibling);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861904.aspx </remarks>
        /// <param name="xML">string xML</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void InsertSubtreeBefore(string xML)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertSubtreeBefore", xML);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864947.aspx </remarks>
        /// <param name="child">NetOffice.OfficeApi.CustomXMLNode child</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void RemoveChild(NetOffice.OfficeApi.CustomXMLNode child)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveChild", child);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862478.aspx </remarks>
        /// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        /// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
        /// <param name="nodeValue">optional string NodeValue = </param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode, object name, object namespaceURI, object nodeType, object nodeValue)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ReplaceChildNode", new object[] { oldNode, name, namespaceURI, nodeType, nodeValue });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862478.aspx </remarks>
        /// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ReplaceChildNode", oldNode);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862478.aspx </remarks>
        /// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
        /// <param name="name">optional string Name = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode, object name)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ReplaceChildNode", oldNode, name);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862478.aspx </remarks>
        /// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode, object name, object namespaceURI)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ReplaceChildNode", oldNode, name, namespaceURI);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862478.aspx </remarks>
        /// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        /// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode, object name, object namespaceURI, object nodeType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ReplaceChildNode", oldNode, name, namespaceURI, nodeType);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863134.aspx </remarks>
        /// <param name="xML">string xML</param>
        /// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ReplaceChildSubtree(string xML, NetOffice.OfficeApi.CustomXMLNode oldNode)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ReplaceChildSubtree", xML, oldNode);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861411.aspx </remarks>
        /// <param name="xPath">string xPath</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLNodes SelectNodes(string xPath)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLNodes>(this, "SelectNodes", typeof(NetOffice.OfficeApi.CustomXMLNodes), xPath);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862077.aspx </remarks>
        /// <param name="xPath">string xPath</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLNode SelectSingleNode(string xPath)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLNode>(this, "SelectSingleNode", typeof(NetOffice.OfficeApi.CustomXMLNode), xPath);
        }

        #endregion

        #pragma warning restore
    }
}
