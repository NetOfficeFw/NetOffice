using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLControlRange 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLControlRange : COMObject
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
                    _type = typeof(IHTMLControlRange);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IHTMLControlRange(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLControlRange(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLControlRange(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLControlRange(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLControlRange(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLControlRange(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLControlRange() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLControlRange(string progId) : base(progId)
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

		#endregion

		#region Methods

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
		/// <param name="item">NetOffice.MSHTMLApi.IHTMLControlElement item</param>
		[SupportByVersion("MSHTML", 4)]
		public void add(NetOffice.MSHTMLApi.IHTMLControlElement item)
		{
			 Factory.ExecuteMethod(this, "add", item);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("MSHTML", 4)]
		public void remove(Int32 index)
		{
			 Factory.ExecuteMethod(this, "remove", index);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement item(Int32 index)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "item", index);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="varargStart">optional object varargStart</param>
		[SupportByVersion("MSHTML", 4)]
		public void scrollIntoView(object varargStart)
		{
			 Factory.ExecuteMethod(this, "scrollIntoView", varargStart);
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

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public NetOffice.MSHTMLApi.IHTMLElement commonParentElement()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLElement>(this, "commonParentElement");
		}

		#endregion

		#pragma warning restore
	}
}
