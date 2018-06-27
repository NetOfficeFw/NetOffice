using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi.Behind
{
    /// <summary>
    /// XMLNode
    /// </summary>
    [SyntaxBypass]
    public class XMLNode_ : COMObject, NetOffice.WordApi.XMLNode_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public XMLNode_() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="dataOnly">optional bool dataOnly</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835819.aspx
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_XML(object dataOnly)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "XML", dataOnly);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Alias for get_XML
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835819.aspx </remarks>
        /// <param name="dataOnly">optional bool dataOnly</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), Redirect("get_XML")]
        public virtual string XML(object dataOnly)
        {
            return get_XML(dataOnly);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="advanced">optional bool advanced</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822315.aspx
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_ValidationErrorText(object advanced)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ValidationErrorText", advanced);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Alias for get_ValidationErrorText
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822315.aspx </remarks>
        /// <param name="advanced">optional bool advanced</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), Redirect("get_ValidationErrorText")]
        public virtual string ValidationErrorText(object advanced)
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
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840965.aspx </remarks>
    [SupportByVersion("Word", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class XMLNode : XMLNode_, NetOffice.WordApi.XMLNode
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
                    _contractType = typeof(NetOffice.WordApi.XMLNode);
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
                    _type = typeof(XMLNode);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public XMLNode() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821132.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual string BaseName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BaseName");
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195202.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Application Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845704.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual Int32 Creator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836034.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840773.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range Range
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Range", typeof(NetOffice.WordApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821675.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
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
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194015.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual string NamespaceURI
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "NamespaceURI");
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835819.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual string XML
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "XML");
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194580.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.XMLNode NextSibling
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNode>(this, "NextSibling", typeof(NetOffice.WordApi.XMLNode));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840520.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.XMLNode PreviousSibling
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNode>(this, "PreviousSibling", typeof(NetOffice.WordApi.XMLNode));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197131.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.XMLNode ParentNode
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNode>(this, "ParentNode", typeof(NetOffice.WordApi.XMLNode));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838342.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.XMLNode FirstChild
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNode>(this, "FirstChild", typeof(NetOffice.WordApi.XMLNode));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834575.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.XMLNode LastChild
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNode>(this, "LastChild", typeof(NetOffice.WordApi.XMLNode));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844845.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document OwnerDocument
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Document>(this, "OwnerDocument", typeof(NetOffice.WordApi.Document));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845491.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdXMLNodeType NodeType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdXMLNodeType>(this, "NodeType");
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821308.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.XMLNodes ChildNodes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNodes>(this, "ChildNodes", typeof(NetOffice.WordApi.XMLNodes));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197252.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.XMLNodes Attributes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNodes>(this, "Attributes", typeof(NetOffice.WordApi.XMLNodes));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845048.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
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
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838711.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual bool HasChildNodes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasChildNodes");
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.XMLChildNodeSuggestions ChildNodeSuggestions
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLChildNodeSuggestions>(this, "ChildNodeSuggestions", typeof(NetOffice.WordApi.XMLChildNodeSuggestions));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837669.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdXMLNodeLevel Level
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdXMLNodeLevel>(this, "Level");
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821280.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdXMLValidationStatus ValidationStatus
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdXMLValidationStatus>(this, "ValidationStatus");
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SmartTag SmartTag
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SmartTag>(this, "SmartTag", typeof(NetOffice.WordApi.SmartTag));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822315.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual string ValidationErrorText
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ValidationErrorText");
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835974.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual string PlaceholderText
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PlaceholderText");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PlaceholderText", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840848.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual string WordOpenXML
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "WordOpenXML");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838108.aspx </remarks>
        /// <param name="xPath">string xPath</param>
        /// <param name="prefixMapping">optional string PrefixMapping = </param>
        /// <param name="fastSearchSkippingTextNodes">optional bool FastSearchSkippingTextNodes = true</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.XMLNode SelectSingleNode(string xPath, object prefixMapping, object fastSearchSkippingTextNodes)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNode>(this, "SelectSingleNode", typeof(NetOffice.WordApi.XMLNode), xPath, prefixMapping, fastSearchSkippingTextNodes);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838108.aspx </remarks>
        /// <param name="xPath">string xPath</param>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.XMLNode SelectSingleNode(string xPath)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNode>(this, "SelectSingleNode", typeof(NetOffice.WordApi.XMLNode), xPath);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838108.aspx </remarks>
        /// <param name="xPath">string xPath</param>
        /// <param name="prefixMapping">optional string PrefixMapping = </param>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.XMLNode SelectSingleNode(string xPath, object prefixMapping)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNode>(this, "SelectSingleNode", typeof(NetOffice.WordApi.XMLNode), xPath, prefixMapping);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835820.aspx </remarks>
        /// <param name="xPath">string xPath</param>
        /// <param name="prefixMapping">optional string PrefixMapping = </param>
        /// <param name="fastSearchSkippingTextNodes">optional bool FastSearchSkippingTextNodes = true</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.XMLNodes SelectNodes(string xPath, object prefixMapping, object fastSearchSkippingTextNodes)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNodes>(this, "SelectNodes", typeof(NetOffice.WordApi.XMLNodes), xPath, prefixMapping, fastSearchSkippingTextNodes);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835820.aspx </remarks>
        /// <param name="xPath">string xPath</param>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.XMLNodes SelectNodes(string xPath)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNodes>(this, "SelectNodes", typeof(NetOffice.WordApi.XMLNodes), xPath);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835820.aspx </remarks>
        /// <param name="xPath">string xPath</param>
        /// <param name="prefixMapping">optional string PrefixMapping = </param>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.XMLNodes SelectNodes(string xPath, object prefixMapping)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLNodes>(this, "SelectNodes", typeof(NetOffice.WordApi.XMLNodes), xPath, prefixMapping);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836057.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void Delete()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821642.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void Copy()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Copy");
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835130.aspx </remarks>
        /// <param name="childElement">NetOffice.WordApi.XMLNode childElement</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void RemoveChild(NetOffice.WordApi.XMLNode childElement)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveChild", childElement);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195671.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void Cut()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Cut");
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191882.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void Validate()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Validate");
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191859.aspx </remarks>
        /// <param name="status">NetOffice.WordApi.Enums.WdXMLValidationStatus status</param>
        /// <param name="errorText">optional object errorText</param>
        /// <param name="clearedAutomatically">optional bool ClearedAutomatically = true</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void SetValidationError(NetOffice.WordApi.Enums.WdXMLValidationStatus status, object errorText, object clearedAutomatically)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetValidationError", status, errorText, clearedAutomatically);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191859.aspx </remarks>
        /// <param name="status">NetOffice.WordApi.Enums.WdXMLValidationStatus status</param>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void SetValidationError(NetOffice.WordApi.Enums.WdXMLValidationStatus status)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetValidationError", status);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191859.aspx </remarks>
        /// <param name="status">NetOffice.WordApi.Enums.WdXMLValidationStatus status</param>
        /// <param name="errorText">optional object errorText</param>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void SetValidationError(NetOffice.WordApi.Enums.WdXMLValidationStatus status, object errorText)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetValidationError", status, errorText);
        }

        #endregion

        #pragma warning restore
    }
}

