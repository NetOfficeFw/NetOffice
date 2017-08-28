using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface DispHTMLAttributeCollection 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class DispHTMLAttributeCollection : COMObject
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
                    _type = typeof(DispHTMLAttributeCollection);
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DispHTMLAttributeCollection(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLAttributeCollection(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLAttributeCollection(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLAttributeCollection(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLAttributeCollection(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLAttributeCollection() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLAttributeCollection(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 length
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "length");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object _newEnum
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "_newEnum");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 ie8_length
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ie8_length");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object constructor
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "constructor");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">optional object name</param>
		[SupportByVersion("MSHTML", 4)]
		public object item(object name)
		{
			return Factory.ExecuteVariantMethodGet(this, "item", name);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public object item()
		{
			return Factory.ExecuteVariantMethodGet(this, "item");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute getNamedItem(string bstrName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute>(this, "getNamedItem", NetOffice.MSHTMLApi.IHTMLDOMAttribute.LateBindingApiWrapperType, bstrName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppNode">NetOffice.MSHTMLApi.IHTMLDOMAttribute ppNode</param>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute setNamedItem(NetOffice.MSHTMLApi.IHTMLDOMAttribute ppNode)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute>(this, "setNamedItem", NetOffice.MSHTMLApi.IHTMLDOMAttribute.LateBindingApiWrapperType, ppNode);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute removeNamedItem(string bstrName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute>(this, "removeNamedItem", NetOffice.MSHTMLApi.IHTMLDOMAttribute.LateBindingApiWrapperType, bstrName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute ie8_getNamedItem(string bstrName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute>(this, "ie8_getNamedItem", NetOffice.MSHTMLApi.IHTMLDOMAttribute.LateBindingApiWrapperType, bstrName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pNodeIn">NetOffice.MSHTMLApi.IHTMLDOMAttribute pNodeIn</param>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute ie8_setNamedItem(NetOffice.MSHTMLApi.IHTMLDOMAttribute pNodeIn)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute>(this, "ie8_setNamedItem", NetOffice.MSHTMLApi.IHTMLDOMAttribute.LateBindingApiWrapperType, pNodeIn);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute ie8_removeNamedItem(string bstrName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute>(this, "ie8_removeNamedItem", NetOffice.MSHTMLApi.IHTMLDOMAttribute.LateBindingApiWrapperType, bstrName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDOMAttribute ie8_item(Int32 index)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLDOMAttribute>(this, "ie8_item", NetOffice.MSHTMLApi.IHTMLDOMAttribute.LateBindingApiWrapperType, index);
		}

		#endregion

		#pragma warning restore
	}
}



