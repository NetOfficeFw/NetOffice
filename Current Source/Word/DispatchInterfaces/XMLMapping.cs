using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using LateBindingApi.Core;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface XMLMapping 
	/// SupportByVersion Word, 12,14
	///</summary>
	[SupportByVersionAttribute("Word", 12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class XMLMapping : COMObject
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
                    _type = typeof(XMLMapping);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLMapping(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLMapping(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLMapping(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLMapping() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLMapping(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public bool IsMapped
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsMapped", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public NetOffice.OfficeApi.CustomXMLPart CustomXMLPart
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomXMLPart", paramsArray);
				NetOffice.OfficeApi.CustomXMLPart newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CustomXMLPart.LateBindingApiWrapperType) as NetOffice.OfficeApi.CustomXMLPart;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public NetOffice.OfficeApi.CustomXMLNode CustomXMLNode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomXMLNode", paramsArray);
				NetOffice.OfficeApi.CustomXMLNode newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType) as NetOffice.OfficeApi.CustomXMLNode;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
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
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public string PrefixMappings
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrefixMappings", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		/// <param name="xPath">string XPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		/// <param name="source">optional NetOffice.OfficeApi.CustomXMLPart Source = 0</param>
		[SupportByVersionAttribute("Word", 12,14)]
		public bool SetMapping(string xPath, string prefixMapping, NetOffice.OfficeApi.CustomXMLPart source)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xPath, prefixMapping, source);
			object returnItem = Invoker.MethodReturn(this, "SetMapping", paramsArray);
			return (bool)returnItem;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		/// <param name="xPath">string XPath</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14)]
		public bool SetMapping(string xPath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xPath);
			object returnItem = Invoker.MethodReturn(this, "SetMapping", paramsArray);
			return (bool)returnItem;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		/// <param name="xPath">string XPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14)]
		public bool SetMapping(string xPath, string prefixMapping)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xPath, prefixMapping);
			object returnItem = Invoker.MethodReturn(this, "SetMapping", paramsArray);
			return (bool)returnItem;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		/// <param name="node">NetOffice.OfficeApi.CustomXMLNode Node</param>
		[SupportByVersionAttribute("Word", 12,14)]
		public bool SetMappingByNode(NetOffice.OfficeApi.CustomXMLNode node)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(node);
			object returnItem = Invoker.MethodReturn(this, "SetMappingByNode", paramsArray);
			return (bool)returnItem;
		}

		#endregion
		#pragma warning restore
	}
}