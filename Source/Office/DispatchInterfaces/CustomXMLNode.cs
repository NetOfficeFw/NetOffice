using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OfficeApi
{
	///<summary>
	/// DispatchInterface CustomXMLNode 
	/// SupportByVersion Office, 12,14,15
	///</summary>
	///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865242.aspx </remarks>
	[SupportByVersionAttribute("Office", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class CustomXMLNode : _IMsoDispObj
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public CustomXMLNode(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomXMLNode(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomXMLNode(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomXMLNode(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomXMLNode(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomXMLNode() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomXMLNode(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864640.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861370.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public NetOffice.OfficeApi.CustomXMLNodes Attributes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Attributes", paramsArray);
				NetOffice.OfficeApi.CustomXMLNodes newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CustomXMLNodes.LateBindingApiWrapperType) as NetOffice.OfficeApi.CustomXMLNodes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862737.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public string BaseName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BaseName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862357.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public NetOffice.OfficeApi.CustomXMLNodes ChildNodes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChildNodes", paramsArray);
				NetOffice.OfficeApi.CustomXMLNodes newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CustomXMLNodes.LateBindingApiWrapperType) as NetOffice.OfficeApi.CustomXMLNodes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863022.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public NetOffice.OfficeApi.CustomXMLNode FirstChild
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FirstChild", paramsArray);
				NetOffice.OfficeApi.CustomXMLNode newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType) as NetOffice.OfficeApi.CustomXMLNode;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864028.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public NetOffice.OfficeApi.CustomXMLNode LastChild
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LastChild", paramsArray);
				NetOffice.OfficeApi.CustomXMLNode newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType) as NetOffice.OfficeApi.CustomXMLNode;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861516.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public string NamespaceURI
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NamespaceURI", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862522.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public NetOffice.OfficeApi.CustomXMLNode NextSibling
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NextSibling", paramsArray);
				NetOffice.OfficeApi.CustomXMLNode newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType) as NetOffice.OfficeApi.CustomXMLNode;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865216.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NodeType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// Get/Set
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862159.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public string NodeValue
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NodeValue", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "NodeValue", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862788.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public object OwnerDocument
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OwnerDocument", paramsArray);
				COMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864973.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public NetOffice.OfficeApi.CustomXMLPart OwnerPart
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OwnerPart", paramsArray);
				NetOffice.OfficeApi.CustomXMLPart newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CustomXMLPart.LateBindingApiWrapperType) as NetOffice.OfficeApi.CustomXMLPart;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861743.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public NetOffice.OfficeApi.CustomXMLNode PreviousSibling
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PreviousSibling", paramsArray);
				NetOffice.OfficeApi.CustomXMLNode newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType) as NetOffice.OfficeApi.CustomXMLNode;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865519.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public NetOffice.OfficeApi.CustomXMLNode ParentNode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ParentNode", paramsArray);
				NetOffice.OfficeApi.CustomXMLNode newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType) as NetOffice.OfficeApi.CustomXMLNode;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// Get/Set
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863358.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public string Text
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Text", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Text", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860871.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public string XPath
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XPath", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860882.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public string XML
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XML", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
		/// <param name="nodeValue">optional string NodeValue = </param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861364.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void AppendChildNode(object name, object namespaceURI, object nodeType, object nodeValue)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, namespaceURI, nodeType, nodeValue);
			Invoker.Method(this, "AppendChildNode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861364.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void AppendChildNode()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AppendChildNode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="name">optional string Name = </param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861364.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void AppendChildNode(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "AppendChildNode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861364.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void AppendChildNode(object name, object namespaceURI)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, namespaceURI);
			Invoker.Method(this, "AppendChildNode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861364.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void AppendChildNode(object name, object namespaceURI, object nodeType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, namespaceURI, nodeType);
			Invoker.Method(this, "AppendChildNode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="xML">string XML</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862169.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void AppendChildSubtree(string xML)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xML);
			Invoker.Method(this, "AppendChildSubtree", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864986.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863303.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public bool HasChildNodes()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "HasChildNodes", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
		/// <param name="nodeValue">optional string NodeValue = </param>
		/// <param name="nextSibling">optional NetOffice.OfficeApi.CustomXMLNode NextSibling = 0</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863860.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void InsertNodeBefore(object name, object namespaceURI, object nodeType, object nodeValue, object nextSibling)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, namespaceURI, nodeType, nodeValue, nextSibling);
			Invoker.Method(this, "InsertNodeBefore", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863860.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void InsertNodeBefore()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InsertNodeBefore", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="name">optional string Name = </param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863860.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void InsertNodeBefore(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "InsertNodeBefore", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863860.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void InsertNodeBefore(object name, object namespaceURI)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, namespaceURI);
			Invoker.Method(this, "InsertNodeBefore", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863860.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void InsertNodeBefore(object name, object namespaceURI, object nodeType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, namespaceURI, nodeType);
			Invoker.Method(this, "InsertNodeBefore", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
		/// <param name="nodeValue">optional string NodeValue = </param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863860.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void InsertNodeBefore(object name, object namespaceURI, object nodeType, object nodeValue)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, namespaceURI, nodeType, nodeValue);
			Invoker.Method(this, "InsertNodeBefore", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="xML">string XML</param>
		/// <param name="nextSibling">optional NetOffice.OfficeApi.CustomXMLNode NextSibling = 0</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861904.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void InsertSubtreeBefore(string xML, object nextSibling)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xML, nextSibling);
			Invoker.Method(this, "InsertSubtreeBefore", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="xML">string XML</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861904.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void InsertSubtreeBefore(string xML)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xML);
			Invoker.Method(this, "InsertSubtreeBefore", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="child">NetOffice.OfficeApi.CustomXMLNode Child</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864947.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void RemoveChild(NetOffice.OfficeApi.CustomXMLNode child)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(child);
			Invoker.Method(this, "RemoveChild", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode OldNode</param>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
		/// <param name="nodeValue">optional string NodeValue = </param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862478.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode, object name, object namespaceURI, object nodeType, object nodeValue)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(oldNode, name, namespaceURI, nodeType, nodeValue);
			Invoker.Method(this, "ReplaceChildNode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode OldNode</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862478.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(oldNode);
			Invoker.Method(this, "ReplaceChildNode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode OldNode</param>
		/// <param name="name">optional string Name = </param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862478.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(oldNode, name);
			Invoker.Method(this, "ReplaceChildNode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode OldNode</param>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862478.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode, object name, object namespaceURI)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(oldNode, name, namespaceURI);
			Invoker.Method(this, "ReplaceChildNode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode OldNode</param>
		/// <param name="name">optional string Name = </param>
		/// <param name="namespaceURI">optional string NamespaceURI = </param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoCustomXMLNodeType NodeType = 1</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862478.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void ReplaceChildNode(NetOffice.OfficeApi.CustomXMLNode oldNode, object name, object namespaceURI, object nodeType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(oldNode, name, namespaceURI, nodeType);
			Invoker.Method(this, "ReplaceChildNode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="xML">string XML</param>
		/// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode OldNode</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863134.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public void ReplaceChildSubtree(string xML, NetOffice.OfficeApi.CustomXMLNode oldNode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xML, oldNode);
			Invoker.Method(this, "ReplaceChildSubtree", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="xPath">string XPath</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861411.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public NetOffice.OfficeApi.CustomXMLNodes SelectNodes(string xPath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xPath);
			object returnItem = Invoker.MethodReturn(this, "SelectNodes", paramsArray);
			NetOffice.OfficeApi.CustomXMLNodes newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.CustomXMLNodes.LateBindingApiWrapperType) as NetOffice.OfficeApi.CustomXMLNodes;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15
		/// </summary>
		/// <param name="xPath">string XPath</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862077.aspx </remarks>
		[SupportByVersionAttribute("Office", 12,14,15)]
		public NetOffice.OfficeApi.CustomXMLNode SelectSingleNode(string xPath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xPath);
			object returnItem = Invoker.MethodReturn(this, "SelectSingleNode", paramsArray);
			NetOffice.OfficeApi.CustomXMLNode newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType) as NetOffice.OfficeApi.CustomXMLNode;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}