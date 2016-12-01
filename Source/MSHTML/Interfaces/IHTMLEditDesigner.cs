using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.MSHTMLApi
{
	///<summary>
	/// Interface IHTMLEditDesigner 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IHTMLEditDesigner : COMObject
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
                    _type = typeof(IHTMLEditDesigner);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLEditDesigner(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditDesigner(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditDesigner(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditDesigner(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditDesigner(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditDesigner() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditDesigner(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="inEvtDispId">Int32 inEvtDispId</param>
		/// <param name="pIEventObj">NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 PreHandleEvent(Int32 inEvtDispId, NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(inEvtDispId, pIEventObj);
			object returnItem = Invoker.MethodReturn(this, "PreHandleEvent", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="inEvtDispId">Int32 inEvtDispId</param>
		/// <param name="pIEventObj">NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 PostHandleEvent(Int32 inEvtDispId, NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(inEvtDispId, pIEventObj);
			object returnItem = Invoker.MethodReturn(this, "PostHandleEvent", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="inEvtDispId">Int32 inEvtDispId</param>
		/// <param name="pIEventObj">NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 TranslateAccelerator(Int32 inEvtDispId, NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(inEvtDispId, pIEventObj);
			object returnItem = Invoker.MethodReturn(this, "TranslateAccelerator", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="inEvtDispId">Int32 inEvtDispId</param>
		/// <param name="pIEventObj">NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 PostEditorEventNotify(Int32 inEvtDispId, NetOffice.MSHTMLApi.IHTMLEventObj pIEventObj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(inEvtDispId, pIEventObj);
			object returnItem = Invoker.MethodReturn(this, "PostEditorEventNotify", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}