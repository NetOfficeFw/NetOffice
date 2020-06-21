using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi.Enums;
using _CustomXMLPartInterop = NetOffice.OfficeApi.DispatchInterfaces.Interop._CustomXMLPart;
using CustomXMLNodeInterop = NetOffice.OfficeApi.DispatchInterfaces.Interop.CustomXMLNode;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface _CustomXMLPart 
	/// SupportByVersion Office, 12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _CustomXMLPart : _IMsoDispObj
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
                    _type = typeof(_CustomXMLPart);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _CustomXMLPart(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _CustomXMLPart(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomXMLPart(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomXMLPart(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomXMLPart(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomXMLPart(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomXMLPart() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CustomXMLPart(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.DocumentElement"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLNode DocumentElement
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLNode>(this, "DocumentElement", NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.Id"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public string Id
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Id");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.NamespaceURI"/> </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.SchemaCollection"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLSchemaCollection SchemaCollection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLSchemaCollection>(this, "SchemaCollection", NetOffice.OfficeApi.CustomXMLSchemaCollection.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "SchemaCollection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.NamespaceManager"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLPrefixMappings NamespaceManager
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLPrefixMappings>(this, "NamespaceManager", NetOffice.OfficeApi.CustomXMLPrefixMappings.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.XML"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public string XML
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "XML");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.Errors"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLValidationErrors Errors
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLValidationErrors>(this, "Errors", NetOffice.OfficeApi.CustomXMLValidationErrors.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.BuiltIn"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool BuiltIn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BuiltIn");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.AddNode"/> </remarks>
		/// <param name="parent">NetOffice.OfficeApi.CustomXMLNode parent</param>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="nextSibling">optional NetOffice.OfficeApi.CustomXMLNode NextSibling = 0</param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
		/// <param name="nodeValue">optional string NodeValue = </param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void AddNode(NetOffice.OfficeApi.CustomXMLNode parent, string name, string namespaceURI, NetOffice.OfficeApi.CustomXMLNode nextSibling, MsoCustomXMLNodeType nodeType, string nodeValue)
		{
			//Factory.ExecuteMethod(this, "AddNode", new object[]{ parent, name, namespaceURI, nextSibling, nodeType, nodeValue });
            _CustomXMLPartInterop proxy = this.UnderlyingObject as _CustomXMLPartInterop;
            if (proxy != null)
            {
                var parentProxy = parent?.UnderlyingObject as CustomXMLNodeInterop;
                var nextSiblingProxy = nextSibling?.UnderlyingObject as CustomXMLNodeInterop;
                proxy.AddNode(parentProxy, name, namespaceURI, nextSiblingProxy, nodeType, nodeValue);
            }
        }

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.AddNode"/> </remarks>
		/// <param name="parent">NetOffice.OfficeApi.CustomXMLNode parent</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void AddNode(NetOffice.OfficeApi.CustomXMLNode parent)
		{
			 Factory.ExecuteMethod(this, "AddNode", parent);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.AddNode"/> </remarks>
		/// <param name="parent">NetOffice.OfficeApi.CustomXMLNode parent</param>
		/// <param name="name">optional string Name = </param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void AddNode(NetOffice.OfficeApi.CustomXMLNode parent, object name)
		{
			 Factory.ExecuteMethod(this, "AddNode", parent, name);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.AddNode"/> </remarks>
		/// <param name="parent">NetOffice.OfficeApi.CustomXMLNode parent</param>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void AddNode(NetOffice.OfficeApi.CustomXMLNode parent, object name, object namespaceURI)
		{
			 Factory.ExecuteMethod(this, "AddNode", parent, name, namespaceURI);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.AddNode"/> </remarks>
		/// <param name="parent">NetOffice.OfficeApi.CustomXMLNode parent</param>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="nextSibling">optional NetOffice.OfficeApi.CustomXMLNode NextSibling = 0</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void AddNode(NetOffice.OfficeApi.CustomXMLNode parent, object name, object namespaceURI, object nextSibling)
		{
			 Factory.ExecuteMethod(this, "AddNode", parent, name, namespaceURI, nextSibling);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.AddNode"/> </remarks>
		/// <param name="parent">NetOffice.OfficeApi.CustomXMLNode parent</param>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="nextSibling">optional NetOffice.OfficeApi.CustomXMLNode NextSibling = 0</param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void AddNode(NetOffice.OfficeApi.CustomXMLNode parent, object name, object namespaceURI, object nextSibling, object nodeType)
		{
			 Factory.ExecuteMethod(this, "AddNode", new object[]{ parent, name, namespaceURI, nextSibling, nodeType });
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.Delete"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.Load"/> </remarks>
		/// <param name="filePath">string filePath</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool Load(string filePath)
		{
			return Factory.ExecuteBoolMethodGet(this, "Load", filePath);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.LoadXML"/> </remarks>
		/// <param name="xML">string xML</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool LoadXML(string xML)
		{
			return Factory.ExecuteBoolMethodGet(this, "LoadXML", xML);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.SelectNodes"/> </remarks>
		/// <param name="xPath">string xPath</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLNodes SelectNodes(string xPath)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CustomXMLNodes>(this, "SelectNodes", NetOffice.OfficeApi.CustomXMLNodes.LateBindingApiWrapperType, xPath);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CustomXMLPart.SelectSingleNode"/> </remarks>
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
