﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLWindow2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLWindow2 : IHTMLFramesCollection2
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
                    _type = typeof(IHTMLWindow2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IHTMLWindow2(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLWindow2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLWindow2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLWindow2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLWindow2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLWindow2(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLWindow2() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLWindow2(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLFramesCollection2 frames
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLFramesCollection2>(this, "frames");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string defaultStatus
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "defaultStatus");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "defaultStatus", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string status
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "status");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "status", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLImageElementFactory Image
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLImageElementFactory>(this, "Image");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLLocation location
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLLocation>(this, "location", NetOffice.MSHTMLApi.IHTMLLocation.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IOmHistory history
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IOmHistory>(this, "history", NetOffice.MSHTMLApi.IOmHistory.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object opener
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "opener");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "opener", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IOmNavigator navigator
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IOmNavigator>(this, "navigator", NetOffice.MSHTMLApi.IOmNavigator.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "name", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLWindow2 parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLWindow2>(this, "parent", NetOffice.MSHTMLApi.IHTMLWindow2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLWindow2 self
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLWindow2>(this, "self", NetOffice.MSHTMLApi.IHTMLWindow2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLWindow2 top
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLWindow2>(this, "top", NetOffice.MSHTMLApi.IHTMLWindow2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLWindow2 window
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLWindow2>(this, "window", NetOffice.MSHTMLApi.IHTMLWindow2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onfocus
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onfocus");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onfocus", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onblur
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onblur");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onblur", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onload
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onload");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onload", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onbeforeunload
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onbeforeunload");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onbeforeunload", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onunload
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onunload");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onunload", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onhelp
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onhelp");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onhelp", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onerror
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onerror");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onerror", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onresize
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onresize");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onresize", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object onscroll
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "onscroll");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "onscroll", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLDocument2 document
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLDocument2>(this, "document", NetOffice.MSHTMLApi.IHTMLDocument2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLEventObj get_event()
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLEventObj>(this, "event", NetOffice.MSHTMLApi.IHTMLEventObj.LateBindingApiWrapperType);
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
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLScreen screen
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLScreen>(this, "screen");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLOptionElementFactory Option
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLOptionElementFactory>(this, "Option");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool closed
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "closed");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IOmNavigator clientInformation
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSHTMLApi.IOmNavigator>(this, "clientInformation", NetOffice.MSHTMLApi.IOmNavigator.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public object offscreenBuffering
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "offscreenBuffering");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "offscreenBuffering", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		public object external
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "external");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="expression">string expression</param>
		/// <param name="msec">Int32 msec</param>
		/// <param name="language">optional object language</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 setTimeout(string expression, Int32 msec, object language)
		{
			return Factory.ExecuteInt32MethodGet(this, "setTimeout", expression, msec, language);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="expression">string expression</param>
		/// <param name="msec">Int32 msec</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public Int32 setTimeout(string expression, Int32 msec)
		{
			return Factory.ExecuteInt32MethodGet(this, "setTimeout", expression, msec);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="timerID">Int32 timerID</param>
		[SupportByVersion("MSHTML", 4)]
		public void clearTimeout(Int32 timerID)
		{
			 Factory.ExecuteMethod(this, "clearTimeout", timerID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="message">optional string message = </param>
		[SupportByVersion("MSHTML", 4)]
		public void alert(object message)
		{
			 Factory.ExecuteMethod(this, "alert", message);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void alert()
		{
			 Factory.ExecuteMethod(this, "alert");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="message">optional string message = </param>
		[SupportByVersion("MSHTML", 4)]
		public bool confirm(object message)
		{
			return Factory.ExecuteBoolMethodGet(this, "confirm", message);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public bool confirm()
		{
			return Factory.ExecuteBoolMethodGet(this, "confirm");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="message">optional string message = </param>
		/// <param name="defstr">optional string defstr = undefined</param>
		[SupportByVersion("MSHTML", 4)]
		public object prompt(object message, object defstr)
		{
			return Factory.ExecuteVariantMethodGet(this, "prompt", message, defstr);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public object prompt()
		{
			return Factory.ExecuteVariantMethodGet(this, "prompt");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="message">optional string message = </param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public object prompt(object message)
		{
			return Factory.ExecuteVariantMethodGet(this, "prompt", message);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void close()
		{
			 Factory.ExecuteMethod(this, "close");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="url">optional string url = </param>
		/// <param name="name">optional string name = </param>
		/// <param name="features">optional string features = </param>
		/// <param name="replace">optional bool replace = false</param>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLWindow2 open(object url, object name, object features, object replace)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLWindow2>(this, "open", NetOffice.MSHTMLApi.IHTMLWindow2.LateBindingApiWrapperType, url, name, features, replace);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLWindow2 open()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLWindow2>(this, "open", NetOffice.MSHTMLApi.IHTMLWindow2.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="url">optional string url = </param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLWindow2 open(object url)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLWindow2>(this, "open", NetOffice.MSHTMLApi.IHTMLWindow2.LateBindingApiWrapperType, url);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="url">optional string url = </param>
		/// <param name="name">optional string name = </param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLWindow2 open(object url, object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLWindow2>(this, "open", NetOffice.MSHTMLApi.IHTMLWindow2.LateBindingApiWrapperType, url, name);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="url">optional string url = </param>
		/// <param name="name">optional string name = </param>
		/// <param name="features">optional string features = </param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLWindow2 open(object url, object name, object features)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLWindow2>(this, "open", NetOffice.MSHTMLApi.IHTMLWindow2.LateBindingApiWrapperType, url, name, features);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="url">string url</param>
		[SupportByVersion("MSHTML", 4)]
		public void navigate(string url)
		{
			 Factory.ExecuteMethod(this, "navigate", url);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="dialog">string dialog</param>
		/// <param name="varArgIn">optional object varArgIn</param>
		/// <param name="varOptions">optional object varOptions</param>
		[SupportByVersion("MSHTML", 4)]
		public object showModalDialog(string dialog, object varArgIn, object varOptions)
		{
			return Factory.ExecuteVariantMethodGet(this, "showModalDialog", dialog, varArgIn, varOptions);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="dialog">string dialog</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public object showModalDialog(string dialog)
		{
			return Factory.ExecuteVariantMethodGet(this, "showModalDialog", dialog);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="dialog">string dialog</param>
		/// <param name="varArgIn">optional object varArgIn</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public object showModalDialog(string dialog, object varArgIn)
		{
			return Factory.ExecuteVariantMethodGet(this, "showModalDialog", dialog, varArgIn);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="helpURL">string helpURL</param>
		/// <param name="helpArg">optional object helpArg</param>
		/// <param name="features">optional string features = </param>
		[SupportByVersion("MSHTML", 4)]
		public void showHelp(string helpURL, object helpArg, object features)
		{
			 Factory.ExecuteMethod(this, "showHelp", helpURL, helpArg, features);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="helpURL">string helpURL</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void showHelp(string helpURL)
		{
			 Factory.ExecuteMethod(this, "showHelp", helpURL);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="helpURL">string helpURL</param>
		/// <param name="helpArg">optional object helpArg</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void showHelp(string helpURL, object helpArg)
		{
			 Factory.ExecuteMethod(this, "showHelp", helpURL, helpArg);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void focus()
		{
			 Factory.ExecuteMethod(this, "focus");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void blur()
		{
			 Factory.ExecuteMethod(this, "blur");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("MSHTML", 4)]
		public void scroll(Int32 x, Int32 y)
		{
			 Factory.ExecuteMethod(this, "scroll", x, y);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="expression">string expression</param>
		/// <param name="msec">Int32 msec</param>
		/// <param name="language">optional object language</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 setInterval(string expression, Int32 msec, object language)
		{
			return Factory.ExecuteInt32MethodGet(this, "setInterval", expression, msec, language);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="expression">string expression</param>
		/// <param name="msec">Int32 msec</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public Int32 setInterval(string expression, Int32 msec)
		{
			return Factory.ExecuteInt32MethodGet(this, "setInterval", expression, msec);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="timerID">Int32 timerID</param>
		[SupportByVersion("MSHTML", 4)]
		public void clearInterval(Int32 timerID)
		{
			 Factory.ExecuteMethod(this, "clearInterval", timerID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="code">string code</param>
		/// <param name="language">optional string language = jScript</param>
		[SupportByVersion("MSHTML", 4)]
		public object execScript(string code, object language)
		{
			return Factory.ExecuteVariantMethodGet(this, "execScript", code, language);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="code">string code</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public object execScript(string code)
		{
			return Factory.ExecuteVariantMethodGet(this, "execScript", code);
		}

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
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("MSHTML", 4)]
		public void scrollBy(Int32 x, Int32 y)
		{
			 Factory.ExecuteMethod(this, "scrollBy", x, y);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("MSHTML", 4)]
		public void scrollTo(Int32 x, Int32 y)
		{
			 Factory.ExecuteMethod(this, "scrollTo", x, y);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("MSHTML", 4)]
		public void moveTo(Int32 x, Int32 y)
		{
			 Factory.ExecuteMethod(this, "moveTo", x, y);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("MSHTML", 4)]
		public void moveBy(Int32 x, Int32 y)
		{
			 Factory.ExecuteMethod(this, "moveBy", x, y);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("MSHTML", 4)]
		public void resizeTo(Int32 x, Int32 y)
		{
			 Factory.ExecuteMethod(this, "resizeTo", x, y);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("MSHTML", 4)]
		public void resizeBy(Int32 x, Int32 y)
		{
			 Factory.ExecuteMethod(this, "resizeBy", x, y);
		}

		#endregion

		#pragma warning restore
	}
}
