using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface _CustomXMLPart 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
    public class _CustomXMLPart : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi._CustomXMLPart
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
                    _contractType = typeof(NetOffice.OfficeApi._CustomXMLPart);
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
                    _type = typeof(_CustomXMLPart);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _CustomXMLPart() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860227.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862841.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLNode DocumentElement
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLNode>(this, "DocumentElement", typeof(NetOffice.OfficeApi.CustomXMLNode));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862405.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string Id
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Id");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860609.aspx </remarks>
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
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862230.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLSchemaCollection SchemaCollection
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLSchemaCollection>(this, "SchemaCollection", typeof(NetOffice.OfficeApi.CustomXMLSchemaCollection));
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "SchemaCollection", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861512.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLPrefixMappings NamespaceManager
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLPrefixMappings>(this, "NamespaceManager", typeof(NetOffice.OfficeApi.CustomXMLPrefixMappings));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860262.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string XML
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "XML");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860307.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLValidationErrors Errors
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLValidationErrors>(this, "Errors", typeof(NetOffice.OfficeApi.CustomXMLValidationErrors));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865215.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool BuiltIn
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "BuiltIn");
            }
        }

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
        public virtual void AddNode(NetOffice.OfficeApi.CustomXMLNode parent, object name, object namespaceURI, object nextSibling, object nodeType, object nodeValue)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AddNode", new object[] { parent, name, namespaceURI, nextSibling, nodeType, nodeValue });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864167.aspx </remarks>
        /// <param name="parent">NetOffice.OfficeApi.CustomXMLNode parent</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void AddNode(NetOffice.OfficeApi.CustomXMLNode parent)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AddNode", parent);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864167.aspx </remarks>
        /// <param name="parent">NetOffice.OfficeApi.CustomXMLNode parent</param>
        /// <param name="name">optional string Name = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void AddNode(NetOffice.OfficeApi.CustomXMLNode parent, object name)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AddNode", parent, name);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864167.aspx </remarks>
        /// <param name="parent">NetOffice.OfficeApi.CustomXMLNode parent</param>
        /// <param name="name">optional string Name = </param>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void AddNode(NetOffice.OfficeApi.CustomXMLNode parent, object name, object namespaceURI)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AddNode", parent, name, namespaceURI);
        }

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
        public virtual void AddNode(NetOffice.OfficeApi.CustomXMLNode parent, object name, object namespaceURI, object nextSibling)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AddNode", parent, name, namespaceURI, nextSibling);
        }

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
        public virtual void AddNode(NetOffice.OfficeApi.CustomXMLNode parent, object name, object namespaceURI, object nextSibling, object nodeType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AddNode", new object[] { parent, name, namespaceURI, nextSibling, nodeType });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861077.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Delete()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865466.aspx </remarks>
        /// <param name="filePath">string filePath</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool Load(string filePath)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Load", filePath);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865293.aspx </remarks>
        /// <param name="xML">string xML</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool LoadXML(string xML)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "LoadXML", xML);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864156.aspx </remarks>
        /// <param name="xPath">string xPath</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLNodes SelectNodes(string xPath)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLNodes>(this, "SelectNodes", typeof(NetOffice.OfficeApi.CustomXMLNodes), xPath);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860915.aspx </remarks>
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
