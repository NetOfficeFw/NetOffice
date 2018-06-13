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
    public interface XMLNode_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="dataOnly">optional bool dataOnly</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835819.aspx
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_XML(object dataOnly);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Alias for get_XML
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835819.aspx </remarks>
        /// <param name="dataOnly">optional bool dataOnly</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), Redirect("get_XML")]
        string XML(object dataOnly);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="advanced">optional bool advanced</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822315.aspx
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_ValidationErrorText(object advanced);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Alias for get_ValidationErrorText
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822315.aspx </remarks>
        /// <param name="advanced">optional bool advanced</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), Redirect("get_ValidationErrorText")]
        string ValidationErrorText(object advanced);

        #endregion
    }

    /// <summary>
    /// DispatchInterface XMLNode 
    /// SupportByVersion Word, 11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840965.aspx </remarks>
    [SupportByVersion("Word", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("09760240-0B89-49F7-A79D-479F24723F56")]
    public interface XMLNode : XMLNode_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821132.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        string BaseName { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195202.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845704.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        Int32 Creator { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836034.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840773.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Range Range { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821675.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        string Text { get; set; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194015.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        string NamespaceURI { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835819.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        new string XML { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194580.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.XMLNode NextSibling { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840520.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.XMLNode PreviousSibling { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197131.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.XMLNode ParentNode { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838342.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.XMLNode FirstChild { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834575.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.XMLNode LastChild { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844845.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Document OwnerDocument { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845491.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdXMLNodeType NodeType { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821308.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.XMLNodes ChildNodes { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197252.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.XMLNodes Attributes { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845048.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        string NodeValue { get; set; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838711.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        bool HasChildNodes { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.XMLChildNodeSuggestions ChildNodeSuggestions { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837669.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdXMLNodeLevel Level { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821280.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdXMLValidationStatus ValidationStatus { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SmartTag SmartTag { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822315.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        new string ValidationErrorText { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835974.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        string PlaceholderText { get; set; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840848.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        string WordOpenXML { get; }

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
        NetOffice.WordApi.XMLNode SelectSingleNode(string xPath, object prefixMapping, object fastSearchSkippingTextNodes);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838108.aspx </remarks>
        /// <param name="xPath">string xPath</param>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.XMLNode SelectSingleNode(string xPath);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838108.aspx </remarks>
        /// <param name="xPath">string xPath</param>
        /// <param name="prefixMapping">optional string PrefixMapping = </param>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.XMLNode SelectSingleNode(string xPath, object prefixMapping);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835820.aspx </remarks>
        /// <param name="xPath">string xPath</param>
        /// <param name="prefixMapping">optional string PrefixMapping = </param>
        /// <param name="fastSearchSkippingTextNodes">optional bool FastSearchSkippingTextNodes = true</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.XMLNodes SelectNodes(string xPath, object prefixMapping, object fastSearchSkippingTextNodes);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835820.aspx </remarks>
        /// <param name="xPath">string xPath</param>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.XMLNodes SelectNodes(string xPath);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835820.aspx </remarks>
        /// <param name="xPath">string xPath</param>
        /// <param name="prefixMapping">optional string PrefixMapping = </param>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.XMLNodes SelectNodes(string xPath, object prefixMapping);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836057.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void Delete();

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821642.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void Copy();

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835130.aspx </remarks>
        /// <param name="childElement">NetOffice.WordApi.XMLNode childElement</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void RemoveChild(NetOffice.WordApi.XMLNode childElement);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195671.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void Cut();

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191882.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void Validate();

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191859.aspx </remarks>
        /// <param name="status">NetOffice.WordApi.Enums.WdXMLValidationStatus status</param>
        /// <param name="errorText">optional object errorText</param>
        /// <param name="clearedAutomatically">optional bool ClearedAutomatically = true</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void SetValidationError(NetOffice.WordApi.Enums.WdXMLValidationStatus status, object errorText, object clearedAutomatically);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191859.aspx </remarks>
        /// <param name="status">NetOffice.WordApi.Enums.WdXMLValidationStatus status</param>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void SetValidationError(NetOffice.WordApi.Enums.WdXMLValidationStatus status);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191859.aspx </remarks>
        /// <param name="status">NetOffice.WordApi.Enums.WdXMLValidationStatus status</param>
        /// <param name="errorText">optional object errorText</param>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void SetValidationError(NetOffice.WordApi.Enums.WdXMLValidationStatus status, object errorText);

        #endregion
    }
}
