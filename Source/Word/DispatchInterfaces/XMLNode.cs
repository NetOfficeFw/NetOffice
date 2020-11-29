using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// XMLNode
	/// </summary>
	[SyntaxBypass]
 	public class XMLNode_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public XMLNode_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public XMLNode_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNode_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNode_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNode_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNode_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNode_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNode_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="dataOnly">optional bool dataOnly</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XML"/>
		[SupportByVersion("Word", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_XML(object dataOnly)
		{
			return Factory.ExecuteStringPropertyGet(this, "XML", dataOnly);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Alias for get_XML
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XML"/> </remarks>
		/// <param name="dataOnly">optional bool dataOnly</param>
		[SupportByVersion("Word", 11,12,14,15,16), Redirect("get_XML")]
		public string XML(object dataOnly)
		{
			return get_XML(dataOnly);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="advanced">optional bool advanced</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.ValidationErrorText"/>
		[SupportByVersion("Word", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_ValidationErrorText(object advanced)
		{
			return Factory.ExecuteStringPropertyGet(this, "ValidationErrorText", advanced);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Alias for get_ValidationErrorText
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.ValidationErrorText"/> </remarks>
		/// <param name="advanced">optional bool advanced</param>
		[SupportByVersion("Word", 11,12,14,15,16), Redirect("get_ValidationErrorText")]
		public string ValidationErrorText(object advanced)
		{
			return get_ValidationErrorText(advanced);
		}

		#endregion

		#region Methods

		#endregion
	}

	/// <summary>
	/// DispatchInterface XMLNode 
	/// SupportByVersion Word, 11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode"/> </remarks>
	[SupportByVersion("Word", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class XMLNode : XMLNode_
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
                    _type = typeof(XMLNode);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public XMLNode(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public XMLNode(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNode(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNode(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNode(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNode(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNode() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLNode(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.BaseName"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public string BaseName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BaseName");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.Application"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.Creator"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.Parent"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.Range"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Range Range
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Range", NetOffice.WordApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.Text"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
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
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.NamespaceURI"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public string NamespaceURI
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NamespaceURI");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XML"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public string XML
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "XML");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.NextSibling"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNode NextSibling
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNode>(this, "NextSibling", NetOffice.WordApi.XMLNode.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.PreviousSibling"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNode PreviousSibling
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNode>(this, "PreviousSibling", NetOffice.WordApi.XMLNode.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.ParentNode"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNode ParentNode
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNode>(this, "ParentNode", NetOffice.WordApi.XMLNode.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.FirstChild"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNode FirstChild
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNode>(this, "FirstChild", NetOffice.WordApi.XMLNode.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.LastChild"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNode LastChild
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNode>(this, "LastChild", NetOffice.WordApi.XMLNode.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.OwnerDocument"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document OwnerDocument
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Document>(this, "OwnerDocument", NetOffice.WordApi.Document.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.NodeType"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdXMLNodeType NodeType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdXMLNodeType>(this, "NodeType");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.ChildNodes"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNodes ChildNodes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNodes>(this, "ChildNodes", NetOffice.WordApi.XMLNodes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.Attributes"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNodes Attributes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNodes>(this, "Attributes", NetOffice.WordApi.XMLNodes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.NodeValue"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
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
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.HasChildNodes"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool HasChildNodes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasChildNodes");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLChildNodeSuggestions ChildNodeSuggestions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLChildNodeSuggestions>(this, "ChildNodeSuggestions", NetOffice.WordApi.XMLChildNodeSuggestions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.Level"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdXMLNodeLevel Level
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdXMLNodeLevel>(this, "Level");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.ValidationStatus"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdXMLValidationStatus ValidationStatus
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdXMLValidationStatus>(this, "ValidationStatus");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.SmartTag SmartTag
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SmartTag>(this, "SmartTag", NetOffice.WordApi.SmartTag.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.ValidationErrorText"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public string ValidationErrorText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ValidationErrorText");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.PlaceholderText"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public string PlaceholderText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PlaceholderText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PlaceholderText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.WordOpenXML"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public string WordOpenXML
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "WordOpenXML");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.SelectSingleNode"/> </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		/// <param name="fastSearchSkippingTextNodes">optional bool FastSearchSkippingTextNodes = true</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNode SelectSingleNode(string xPath, object prefixMapping, object fastSearchSkippingTextNodes)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNode>(this, "SelectSingleNode", NetOffice.WordApi.XMLNode.LateBindingApiWrapperType, xPath, prefixMapping, fastSearchSkippingTextNodes);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.SelectSingleNode"/> </remarks>
		/// <param name="xPath">string xPath</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNode SelectSingleNode(string xPath)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNode>(this, "SelectSingleNode", NetOffice.WordApi.XMLNode.LateBindingApiWrapperType, xPath);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.SelectSingleNode"/> </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNode SelectSingleNode(string xPath, object prefixMapping)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNode>(this, "SelectSingleNode", NetOffice.WordApi.XMLNode.LateBindingApiWrapperType, xPath, prefixMapping);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.SelectNodes"/> </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		/// <param name="fastSearchSkippingTextNodes">optional bool FastSearchSkippingTextNodes = true</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNodes SelectNodes(string xPath, object prefixMapping, object fastSearchSkippingTextNodes)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNodes>(this, "SelectNodes", NetOffice.WordApi.XMLNodes.LateBindingApiWrapperType, xPath, prefixMapping, fastSearchSkippingTextNodes);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.SelectNodes"/> </remarks>
		/// <param name="xPath">string xPath</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNodes SelectNodes(string xPath)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNodes>(this, "SelectNodes", NetOffice.WordApi.XMLNodes.LateBindingApiWrapperType, xPath);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.SelectNodes"/> </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNodes SelectNodes(string xPath, object prefixMapping)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNodes>(this, "SelectNodes", NetOffice.WordApi.XMLNodes.LateBindingApiWrapperType, xPath, prefixMapping);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.Delete"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.Copy"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.RemoveChild"/> </remarks>
		/// <param name="childElement">NetOffice.WordApi.XMLNode childElement</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void RemoveChild(NetOffice.WordApi.XMLNode childElement)
		{
			 Factory.ExecuteMethod(this, "RemoveChild", childElement);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.Cut"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void Cut()
		{
			 Factory.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.Validate"/> </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void Validate()
		{
			 Factory.ExecuteMethod(this, "Validate");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.SetValidationError"/> </remarks>
		/// <param name="status">NetOffice.WordApi.Enums.WdXMLValidationStatus status</param>
		/// <param name="errorText">optional object errorText</param>
		/// <param name="clearedAutomatically">optional bool ClearedAutomatically = true</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void SetValidationError(NetOffice.WordApi.Enums.WdXMLValidationStatus status, object errorText, object clearedAutomatically)
		{
			 Factory.ExecuteMethod(this, "SetValidationError", status, errorText, clearedAutomatically);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.SetValidationError"/> </remarks>
		/// <param name="status">NetOffice.WordApi.Enums.WdXMLValidationStatus status</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void SetValidationError(NetOffice.WordApi.Enums.WdXMLValidationStatus status)
		{
			 Factory.ExecuteMethod(this, "SetValidationError", status);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.XMLNode.SetValidationError"/> </remarks>
		/// <param name="status">NetOffice.WordApi.Enums.WdXMLValidationStatus status</param>
		/// <param name="errorText">optional object errorText</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void SetValidationError(NetOffice.WordApi.Enums.WdXMLValidationStatus status, object errorText)
		{
			 Factory.ExecuteMethod(this, "SetValidationError", status, errorText);
		}

		#endregion

		#pragma warning restore
	}
}
