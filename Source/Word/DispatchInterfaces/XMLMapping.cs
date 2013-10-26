using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface XMLMapping 
	/// SupportByVersion Word, 12,14,15
	///</summary>
	///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838482.aspx </remarks>
	[SupportByVersionAttribute("Word", 12,14,15)]
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

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public XMLMapping(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLMapping(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLMapping(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XMLMapping(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
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
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838317.aspx </remarks>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836992.aspx </remarks>
		[SupportByVersionAttribute("Word", 12,14,15)]
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
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191939.aspx </remarks>
		[SupportByVersionAttribute("Word", 12,14,15)]
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
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839742.aspx </remarks>
		[SupportByVersionAttribute("Word", 12,14,15)]
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
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836066.aspx </remarks>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public NetOffice.OfficeApi.CustomXMLPart CustomXMLPart
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomXMLPart", paramsArray);
				NetOffice.OfficeApi.CustomXMLPart newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CustomXMLPart.LateBindingApiWrapperType) as NetOffice.OfficeApi.CustomXMLPart;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837855.aspx </remarks>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public NetOffice.OfficeApi.CustomXMLNode CustomXMLNode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomXMLNode", paramsArray);
				NetOffice.OfficeApi.CustomXMLNode newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType) as NetOffice.OfficeApi.CustomXMLNode;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845740.aspx </remarks>
		[SupportByVersionAttribute("Word", 12,14,15)]
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
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845123.aspx </remarks>
		[SupportByVersionAttribute("Word", 12,14,15)]
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
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="xPath">string XPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		/// <param name="source">optional NetOffice.OfficeApi.CustomXMLPart Source = 0</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845439.aspx </remarks>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public bool SetMapping(string xPath, object prefixMapping, object source)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xPath, prefixMapping, source);
			object returnItem = Invoker.MethodReturn(this, "SetMapping", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="xPath">string XPath</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845439.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15)]
		public bool SetMapping(string xPath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xPath);
			object returnItem = Invoker.MethodReturn(this, "SetMapping", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="xPath">string XPath</param>
		/// <param name="prefixMapping">optional string PrefixMapping = </param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845439.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15)]
		public bool SetMapping(string xPath, object prefixMapping)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xPath, prefixMapping);
			object returnItem = Invoker.MethodReturn(this, "SetMapping", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820763.aspx </remarks>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		/// <param name="node">NetOffice.OfficeApi.CustomXMLNode Node</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822994.aspx </remarks>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public bool SetMappingByNode(NetOffice.OfficeApi.CustomXMLNode node)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(node);
			object returnItem = Invoker.MethodReturn(this, "SetMappingByNode", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}