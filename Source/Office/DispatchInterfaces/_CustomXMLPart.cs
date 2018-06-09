using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface _CustomXMLPart 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000CDB05-0000-0000-C000-000000000046")]
    public interface _CustomXMLPart : _IMsoDispObj
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860227.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862841.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLNode DocumentElement { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862405.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        string Id { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860609.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        string NamespaceURI { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862230.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLSchemaCollection SchemaCollection { get; set; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861512.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLPrefixMappings NamespaceManager { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860262.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        string XML { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860307.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLValidationErrors Errors { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865215.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool BuiltIn { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864167.aspx </remarks>
        /// <param name="parent">NetOffice.OfficeApi.CustomXMLNode parent</param>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        /// <param name="nextSibling">optional NetOffice.OfficeApi.CustomXMLNode NextSibling = 0</param>
        /// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
        /// <param name="nodeValue">optional string NodeValue = </param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void AddNode(NetOffice.OfficeApi.CustomXMLNode parent, object name, object namespaceURI, object nextSibling, object nodeType, object nodeValue);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864167.aspx </remarks>
        /// <param name="parent">NetOffice.OfficeApi.CustomXMLNode parent</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void AddNode(NetOffice.OfficeApi.CustomXMLNode parent);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864167.aspx </remarks>
        /// <param name="parent">NetOffice.OfficeApi.CustomXMLNode parent</param>
        /// <param name="name">optional string Name = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void AddNode(NetOffice.OfficeApi.CustomXMLNode parent, object name);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864167.aspx </remarks>
        /// <param name="parent">NetOffice.OfficeApi.CustomXMLNode parent</param>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void AddNode(NetOffice.OfficeApi.CustomXMLNode parent, object name, object namespaceURI);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864167.aspx </remarks>
        /// <param name="parent">NetOffice.OfficeApi.CustomXMLNode parent</param>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        /// <param name="nextSibling">optional NetOffice.OfficeApi.CustomXMLNode NextSibling = 0</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void AddNode(NetOffice.OfficeApi.CustomXMLNode parent, object name, object namespaceURI, object nextSibling);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864167.aspx </remarks>
        /// <param name="parent">NetOffice.OfficeApi.CustomXMLNode parent</param>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        /// <param name="nextSibling">optional NetOffice.OfficeApi.CustomXMLNode NextSibling = 0</param>
        /// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void AddNode(NetOffice.OfficeApi.CustomXMLNode parent, object name, object namespaceURI, object nextSibling, object nodeType);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861077.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Delete();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865466.aspx </remarks>
        /// <param name="filePath">string filePath</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
         bool Load(string filePath);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865293.aspx </remarks>
        /// <param name="xML">string xML</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool LoadXML(string xML);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864156.aspx </remarks>
        /// <param name="xPath">string xPath</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLNodes SelectNodes(string xPath);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860915.aspx </remarks>
        /// <param name="xPath">string xPath</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLNode SelectSingleNode(string xPath);

        #endregion
    }
}
