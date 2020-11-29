﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface DispHTMLElementCollection 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class DispHTMLElementCollection : COMObject
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
                    _type = typeof(DispHTMLElementCollection);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public DispHTMLElementCollection(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DispHTMLElementCollection(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLElementCollection(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLElementCollection(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLElementCollection(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLElementCollection(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLElementCollection() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DispHTMLElementCollection(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 length
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "length");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "length", value);
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
		[SupportByVersion("MSHTML", 4)]
		public string toString()
		{
			return Factory.ExecuteStringMethodGet(this, "toString");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="index">optional object index</param>
		[SupportByVersion("MSHTML", 4)]
		public object item(object name, object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "item", name, index);
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
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public object item(object name)
		{
			return Factory.ExecuteVariantMethodGet(this, "item", name);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="tagName">object tagName</param>
		[SupportByVersion("MSHTML", 4)]
		public object tags(object tagName)
		{
			return Factory.ExecuteVariantMethodGet(this, "tags", tagName);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="urn">object urn</param>
		[SupportByVersion("MSHTML", 4)]
		public object urns(object urn)
		{
			return Factory.ExecuteVariantMethodGet(this, "urns", urn);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("MSHTML", 4)]
		public object namedItem(string name)
		{
			return Factory.ExecuteVariantMethodGet(this, "namedItem", name);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement2 ie8_item(Int32 index)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement2>(this, "ie8_item", index);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement2 ie8_namedItem(string name)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement2>(this, "ie8_namedItem", name);
		}

		#endregion

		#pragma warning restore
	}
}
