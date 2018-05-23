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
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865242.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface CustomXMLNode : _IMsoDispObj
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864640.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861370.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLNodes Attributes { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862737.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        string BaseName { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862357.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLNodes ChildNodes { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863022.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLNode FirstChild { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864028.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLNode LastChild { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861516.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        string NamespaceURI { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862522.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLNode NextSibling { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865216.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862159.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        string NodeValue { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862788.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        object OwnerDocument { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864973.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLPart OwnerPart { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861743.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLNode PreviousSibling { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865519.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLNode ParentNode { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863358.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        string Text { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860871.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        string XPath { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860882.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        string XML { get; }

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
        void AppendChildNode(object name, object namespaceURI, object nodeType, object nodeValue);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861364.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void AppendChildNode();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861364.aspx </remarks>
        /// <param name="name">optional string Name = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void AppendChildNode(object name);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861364.aspx </remarks>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void AppendChildNode(object name, object namespaceURI);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861364.aspx </remarks>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        /// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void AppendChildNode(object name, object namespaceURI, object nodeType);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862169.aspx </remarks>
        /// <param name="xML">string xML</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void AppendChildSubtree(string xML);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864986.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Delete();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863303.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool HasChildNodes();

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
        void InsertNodeBefore(object name, object namespaceURI, object nodeType, object nodeValue, object nextSibling);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863860.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void InsertNodeBefore();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863860.aspx </remarks>
        /// <param name="name">optional string Name = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void InsertNodeBefore(object name);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863860.aspx </remarks>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void InsertNodeBefore(object name, object namespaceURI);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863860.aspx </remarks>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        /// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void InsertNodeBefore(object name, object namespaceURI, object nodeType);

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
        void InsertNodeBefore(object name, object namespaceURI, object nodeType, object nodeValue);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861904.aspx </remarks>
        /// <param name="xML">string xML</param>
        /// <param name="nextSibling">optional NetOffice.OfficeApi.CustomXMLNode NextSibling = 0</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void InsertSubtreeBefore(string xML, object nextSibling);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861904.aspx </remarks>
        /// <param name="xML">string xML</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void InsertSubtreeBefore(string xML);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864947.aspx </remarks>
        /// <param name="child">NetOffice.OfficeApi.CustomXMLNode child</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void RemoveChild(NetOffice.OfficeApi.CustomXMLNode child);

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
        void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode, object name, object namespaceURI, object nodeType, object nodeValue);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862478.aspx </remarks>
        /// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862478.aspx </remarks>
        /// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
        /// <param name="name">optional string Name = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode, object name);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862478.aspx </remarks>
        /// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode, object name, object namespaceURI);

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
        void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode, object name, object namespaceURI, object nodeType);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863134.aspx </remarks>
        /// <param name="xML">string xML</param>
        /// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ReplaceChildSubtree(string xML, NetOffice.OfficeApi.CustomXMLNode oldNode);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861411.aspx </remarks>
        /// <param name="xPath">string xPath</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLNodes SelectNodes(string xPath);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862077.aspx </remarks>
        /// <param name="xPath">string xPath</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLNode SelectSingleNode(string xPath);

        #endregion
    }
}
