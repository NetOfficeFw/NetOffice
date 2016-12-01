using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.MSHTMLApi
{
	///<summary>
	/// DispatchInterface IHTMLTxtRange 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class IHTMLTxtRange : COMObject
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
                    _type = typeof(IHTMLTxtRange);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		[SupportByVersionAttribute("MSHTML", 4)]
		public string htmlText
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "htmlText", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public string text
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "text", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "text", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLElement parentElement()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "parentElement", paramsArray);
			NetOffice.MSHTMLApi.IHTMLElement newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSHTMLApi.IHTMLElement;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public NetOffice.MSHTMLApi.IHTMLTxtRange duplicate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "duplicate", paramsArray);
			NetOffice.MSHTMLApi.IHTMLTxtRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSHTMLApi.IHTMLTxtRange.LateBindingApiWrapperType) as NetOffice.MSHTMLApi.IHTMLTxtRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="range">NetOffice.MSHTMLApi.IHTMLTxtRange range</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool inRange(NetOffice.MSHTMLApi.IHTMLTxtRange range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range);
			object returnItem = Invoker.MethodReturn(this, "inRange", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="range">NetOffice.MSHTMLApi.IHTMLTxtRange range</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool isEqual(NetOffice.MSHTMLApi.IHTMLTxtRange range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range);
			object returnItem = Invoker.MethodReturn(this, "isEqual", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="fStart">optional bool fStart = true</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void scrollIntoView(object fStart)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fStart);
			Invoker.Method(this, "scrollIntoView", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public void scrollIntoView()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "scrollIntoView", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="start">optional bool Start = true</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void collapse(object start)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start);
			Invoker.Method(this, "collapse", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public void collapse()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "collapse", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="unit">string Unit</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool expand(string unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "expand", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="unit">string Unit</param>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 move(string unit, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count);
			object returnItem = Invoker.MethodReturn(this, "move", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="unit">string Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 move(string unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "move", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="unit">string Unit</param>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 moveStart(string unit, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count);
			object returnItem = Invoker.MethodReturn(this, "moveStart", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="unit">string Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 moveStart(string unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "moveStart", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="unit">string Unit</param>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 moveEnd(string unit, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit, count);
			object returnItem = Invoker.MethodReturn(this, "moveEnd", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="unit">string Unit</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 moveEnd(string unit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unit);
			object returnItem = Invoker.MethodReturn(this, "moveEnd", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void select()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "select", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="html">string html</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void pasteHTML(string html)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(html);
			Invoker.Method(this, "pasteHTML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="element">NetOffice.MSHTMLApi.IHTMLElement element</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void moveToElementText(NetOffice.MSHTMLApi.IHTMLElement element)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(element);
			Invoker.Method(this, "moveToElementText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="how">string how</param>
		/// <param name="sourceRange">NetOffice.MSHTMLApi.IHTMLTxtRange SourceRange</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void setEndPoint(string how, NetOffice.MSHTMLApi.IHTMLTxtRange sourceRange)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(how, sourceRange);
			Invoker.Method(this, "setEndPoint", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="how">string how</param>
		/// <param name="sourceRange">NetOffice.MSHTMLApi.IHTMLTxtRange SourceRange</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 compareEndPoints(string how, NetOffice.MSHTMLApi.IHTMLTxtRange sourceRange)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(how, sourceRange);
			object returnItem = Invoker.MethodReturn(this, "compareEndPoints", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="_string">string String</param>
		/// <param name="count">optional Int32 Count = 1073741823</param>
		/// <param name="flags">optional Int32 Flags = 0</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool findText(string _string, object count, object flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(_string, count, flags);
			object returnItem = Invoker.MethodReturn(this, "findText", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="_string">string String</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool findText(string _string)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(_string);
			object returnItem = Invoker.MethodReturn(this, "findText", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="_string">string String</param>
		/// <param name="count">optional Int32 Count = 1073741823</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool findText(string _string, object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(_string, count);
			object returnItem = Invoker.MethodReturn(this, "findText", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void moveToPoint(Int32 x, Int32 y)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(x, y);
			Invoker.Method(this, "moveToPoint", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public string getBookmark()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "getBookmark", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="bookmark">string Bookmark</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool moveToBookmark(string bookmark)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bookmark);
			object returnItem = Invoker.MethodReturn(this, "moveToBookmark", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool queryCommandSupported(string cmdID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cmdID);
			object returnItem = Invoker.MethodReturn(this, "queryCommandSupported", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool queryCommandEnabled(string cmdID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cmdID);
			object returnItem = Invoker.MethodReturn(this, "queryCommandEnabled", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool queryCommandState(string cmdID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cmdID);
			object returnItem = Invoker.MethodReturn(this, "queryCommandState", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool queryCommandIndeterm(string cmdID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cmdID);
			object returnItem = Invoker.MethodReturn(this, "queryCommandIndeterm", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public string queryCommandText(string cmdID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cmdID);
			object returnItem = Invoker.MethodReturn(this, "queryCommandText", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public object queryCommandValue(string cmdID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cmdID);
			object returnItem = Invoker.MethodReturn(this, "queryCommandValue", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		/// <param name="showUI">optional bool showUI = false</param>
		/// <param name="value">optional object value</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool execCommand(string cmdID, object showUI, object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cmdID, showUI, value);
			object returnItem = Invoker.MethodReturn(this, "execCommand", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool execCommand(string cmdID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cmdID);
			object returnItem = Invoker.MethodReturn(this, "execCommand", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		/// <param name="showUI">optional bool showUI = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool execCommand(string cmdID, object showUI)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cmdID, showUI);
			object returnItem = Invoker.MethodReturn(this, "execCommand", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public bool execCommandShowHelp(string cmdID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cmdID);
			object returnItem = Invoker.MethodReturn(this, "execCommandShowHelp", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}