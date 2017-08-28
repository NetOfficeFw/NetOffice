using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLTxtRange 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLTxtRange : COMObject
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
                    _type = typeof(IHTMLTxtRange);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IHTMLTxtRange(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLTxtRange(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLTxtRange(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLTxtRange(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLTxtRange(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLTxtRange(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLTxtRange() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLTxtRange(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string htmlText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "htmlText");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string text
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "text");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "text", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement parentElement()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "parentElement");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLTxtRange duplicate()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLTxtRange>(this, "duplicate", NetOffice.MSHTMLApi.IHTMLTxtRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="range">NetOffice.MSHTMLApi.IHTMLTxtRange range</param>
		[SupportByVersion("MSHTML", 4)]
		public bool inRange(NetOffice.MSHTMLApi.IHTMLTxtRange range)
		{
			return Factory.ExecuteBoolMethodGet(this, "inRange", range);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="range">NetOffice.MSHTMLApi.IHTMLTxtRange range</param>
		[SupportByVersion("MSHTML", 4)]
		public bool isEqual(NetOffice.MSHTMLApi.IHTMLTxtRange range)
		{
			return Factory.ExecuteBoolMethodGet(this, "isEqual", range);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fStart">optional bool fStart = true</param>
		[SupportByVersion("MSHTML", 4)]
		public void scrollIntoView(object fStart)
		{
			 Factory.ExecuteMethod(this, "scrollIntoView", fStart);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void scrollIntoView()
		{
			 Factory.ExecuteMethod(this, "scrollIntoView");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="start">optional bool Start = true</param>
		[SupportByVersion("MSHTML", 4)]
		public void collapse(object start)
		{
			 Factory.ExecuteMethod(this, "collapse", start);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void collapse()
		{
			 Factory.ExecuteMethod(this, "collapse");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		[SupportByVersion("MSHTML", 4)]
		public bool expand(string unit)
		{
			return Factory.ExecuteBoolMethodGet(this, "expand", unit);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 move(string unit, object count)
		{
			return Factory.ExecuteInt32MethodGet(this, "move", unit, count);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public Int32 move(string unit)
		{
			return Factory.ExecuteInt32MethodGet(this, "move", unit);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 moveStart(string unit, object count)
		{
			return Factory.ExecuteInt32MethodGet(this, "moveStart", unit, count);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public Int32 moveStart(string unit)
		{
			return Factory.ExecuteInt32MethodGet(this, "moveStart", unit);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 moveEnd(string unit, object count)
		{
			return Factory.ExecuteInt32MethodGet(this, "moveEnd", unit, count);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public Int32 moveEnd(string unit)
		{
			return Factory.ExecuteInt32MethodGet(this, "moveEnd", unit);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void select()
		{
			 Factory.ExecuteMethod(this, "select");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="html">string html</param>
		[SupportByVersion("MSHTML", 4)]
		public void pasteHTML(string html)
		{
			 Factory.ExecuteMethod(this, "pasteHTML", html);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="element">NetOffice.MSHTMLApi.IHTMLElement element</param>
		[SupportByVersion("MSHTML", 4)]
		public void moveToElementText(NetOffice.MSHTMLApi.IHTMLElement element)
		{
			 Factory.ExecuteMethod(this, "moveToElementText", element);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="how">string how</param>
		/// <param name="sourceRange">NetOffice.MSHTMLApi.IHTMLTxtRange sourceRange</param>
		[SupportByVersion("MSHTML", 4)]
		public void setEndPoint(string how, NetOffice.MSHTMLApi.IHTMLTxtRange sourceRange)
		{
			 Factory.ExecuteMethod(this, "setEndPoint", how, sourceRange);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="how">string how</param>
		/// <param name="sourceRange">NetOffice.MSHTMLApi.IHTMLTxtRange sourceRange</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 compareEndPoints(string how, NetOffice.MSHTMLApi.IHTMLTxtRange sourceRange)
		{
			return Factory.ExecuteInt32MethodGet(this, "compareEndPoints", how, sourceRange);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="_string">string string</param>
		/// <param name="count">optional Int32 Count = 1073741823</param>
		/// <param name="flags">optional Int32 Flags = 0</param>
		[SupportByVersion("MSHTML", 4)]
		public bool findText(string _string, object count, object flags)
		{
			return Factory.ExecuteBoolMethodGet(this, "findText", _string, count, flags);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="_string">string string</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public bool findText(string _string)
		{
			return Factory.ExecuteBoolMethodGet(this, "findText", _string);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="_string">string string</param>
		/// <param name="count">optional Int32 Count = 1073741823</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public bool findText(string _string, object count)
		{
			return Factory.ExecuteBoolMethodGet(this, "findText", _string, count);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("MSHTML", 4)]
		public void moveToPoint(Int32 x, Int32 y)
		{
			 Factory.ExecuteMethod(this, "moveToPoint", x, y);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public string getBookmark()
		{
			return Factory.ExecuteStringMethodGet(this, "getBookmark");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bookmark">string bookmark</param>
		[SupportByVersion("MSHTML", 4)]
		public bool moveToBookmark(string bookmark)
		{
			return Factory.ExecuteBoolMethodGet(this, "moveToBookmark", bookmark);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public bool queryCommandSupported(string cmdID)
		{
			return Factory.ExecuteBoolMethodGet(this, "queryCommandSupported", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public bool queryCommandEnabled(string cmdID)
		{
			return Factory.ExecuteBoolMethodGet(this, "queryCommandEnabled", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public bool queryCommandState(string cmdID)
		{
			return Factory.ExecuteBoolMethodGet(this, "queryCommandState", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public bool queryCommandIndeterm(string cmdID)
		{
			return Factory.ExecuteBoolMethodGet(this, "queryCommandIndeterm", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public string queryCommandText(string cmdID)
		{
			return Factory.ExecuteStringMethodGet(this, "queryCommandText", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public object queryCommandValue(string cmdID)
		{
			return Factory.ExecuteVariantMethodGet(this, "queryCommandValue", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		/// <param name="showUI">optional bool showUI = false</param>
		/// <param name="value">optional object value</param>
		[SupportByVersion("MSHTML", 4)]
		public bool execCommand(string cmdID, object showUI, object value)
		{
			return Factory.ExecuteBoolMethodGet(this, "execCommand", cmdID, showUI, value);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public bool execCommand(string cmdID)
		{
			return Factory.ExecuteBoolMethodGet(this, "execCommand", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		/// <param name="showUI">optional bool showUI = false</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public bool execCommand(string cmdID, object showUI)
		{
			return Factory.ExecuteBoolMethodGet(this, "execCommand", cmdID, showUI);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public bool execCommandShowHelp(string cmdID)
		{
			return Factory.ExecuteBoolMethodGet(this, "execCommandShowHelp", cmdID);
		}

		#endregion

		#pragma warning restore
	}
}
