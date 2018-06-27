using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLTxtRange 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLTxtRange : COMObject, NetOffice.MSHTMLApi.IHTMLTxtRange
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLTxtRange);
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
                    _type = typeof(IHTMLTxtRange);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLTxtRange() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string htmlText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "htmlText");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string text
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "text");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "text", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLElement parentElement()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "parentElement");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLTxtRange duplicate()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLTxtRange>(this, "duplicate", typeof(NetOffice.MSHTMLApi.IHTMLTxtRange));
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="range">NetOffice.MSHTMLApi.IHTMLTxtRange range</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool inRange(NetOffice.MSHTMLApi.IHTMLTxtRange range)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "inRange", range);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="range">NetOffice.MSHTMLApi.IHTMLTxtRange range</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool isEqual(NetOffice.MSHTMLApi.IHTMLTxtRange range)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "isEqual", range);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fStart">optional bool fStart = true</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void scrollIntoView(object fStart)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "scrollIntoView", fStart);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void scrollIntoView()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "scrollIntoView");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="start">optional bool Start = true</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void collapse(object start)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "collapse", start);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void collapse()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "collapse");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool expand(string unit)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "expand", unit);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 move(string unit, object count)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "move", unit, count);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 move(string unit)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "move", unit);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 moveStart(string unit, object count)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "moveStart", unit, count);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 moveStart(string unit)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "moveStart", unit);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 moveEnd(string unit, object count)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "moveEnd", unit, count);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="unit">string unit</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 moveEnd(string unit)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "moveEnd", unit);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void select()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "select");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="html">string html</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void pasteHTML(string html)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "pasteHTML", html);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="element">NetOffice.MSHTMLApi.IHTMLElement element</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void moveToElementText(NetOffice.MSHTMLApi.IHTMLElement element)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "moveToElementText", element);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="how">string how</param>
		/// <param name="sourceRange">NetOffice.MSHTMLApi.IHTMLTxtRange sourceRange</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void setEndPoint(string how, NetOffice.MSHTMLApi.IHTMLTxtRange sourceRange)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "setEndPoint", how, sourceRange);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="how">string how</param>
		/// <param name="sourceRange">NetOffice.MSHTMLApi.IHTMLTxtRange sourceRange</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 compareEndPoints(string how, NetOffice.MSHTMLApi.IHTMLTxtRange sourceRange)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "compareEndPoints", how, sourceRange);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="_string">string string</param>
		/// <param name="count">optional Int32 Count = 1073741823</param>
		/// <param name="flags">optional Int32 Flags = 0</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool findText(string _string, object count, object flags)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "findText", _string, count, flags);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="_string">string string</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual bool findText(string _string)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "findText", _string);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="_string">string string</param>
		/// <param name="count">optional Int32 Count = 1073741823</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual bool findText(string _string, object count)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "findText", _string, count);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void moveToPoint(Int32 x, Int32 y)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "moveToPoint", x, y);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string getBookmark()
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "getBookmark");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bookmark">string bookmark</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool moveToBookmark(string bookmark)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "moveToBookmark", bookmark);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool queryCommandSupported(string cmdID)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "queryCommandSupported", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool queryCommandEnabled(string cmdID)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "queryCommandEnabled", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool queryCommandState(string cmdID)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "queryCommandState", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool queryCommandIndeterm(string cmdID)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "queryCommandIndeterm", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual string queryCommandText(string cmdID)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "queryCommandText", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object queryCommandValue(string cmdID)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "queryCommandValue", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		/// <param name="showUI">optional bool showUI = false</param>
		/// <param name="value">optional object value</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool execCommand(string cmdID, object showUI, object value)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "execCommand", cmdID, showUI, value);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual bool execCommand(string cmdID)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "execCommand", cmdID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		/// <param name="showUI">optional bool showUI = false</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual bool execCommand(string cmdID, object showUI)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "execCommand", cmdID, showUI);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="cmdID">string cmdID</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool execCommandShowHelp(string cmdID)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "execCommandShowHelp", cmdID);
		}

		#endregion

		#pragma warning restore
	}
}


