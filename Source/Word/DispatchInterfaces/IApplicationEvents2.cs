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
	/// DispatchInterface IApplicationEvents2 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class IApplicationEvents2 : COMObject
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
                    _type = typeof(IApplicationEvents2);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IApplicationEvents2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents2(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents2() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents2(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Startup()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Startup", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Quit()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Quit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void DocumentChange()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DocumentChange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void DocumentOpen(NetOffice.WordApi.Document doc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc);
			Invoker.Method(this, "DocumentOpen", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void DocumentBeforeClose(NetOffice.WordApi.Document doc, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, cancel);
			Invoker.Method(this, "DocumentBeforeClose", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void DocumentBeforePrint(NetOffice.WordApi.Document doc, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, cancel);
			Invoker.Method(this, "DocumentBeforePrint", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		/// <param name="saveAsUI">bool SaveAsUI</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void DocumentBeforeSave(NetOffice.WordApi.Document doc, bool saveAsUI, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, saveAsUI, cancel);
			Invoker.Method(this, "DocumentBeforeSave", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void NewDocument(NetOffice.WordApi.Document doc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc);
			Invoker.Method(this, "NewDocument", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		/// <param name="wn">NetOffice.WordApi.Window Wn</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void WindowActivate(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, wn);
			Invoker.Method(this, "WindowActivate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		/// <param name="wn">NetOffice.WordApi.Window Wn</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void WindowDeactivate(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, wn);
			Invoker.Method(this, "WindowDeactivate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection Sel</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void WindowSelectionChange(NetOffice.WordApi.Selection sel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sel);
			Invoker.Method(this, "WindowSelectionChange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection Sel</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void WindowBeforeRightClick(NetOffice.WordApi.Selection sel, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sel, cancel);
			Invoker.Method(this, "WindowBeforeRightClick", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection Sel</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void WindowBeforeDoubleClick(NetOffice.WordApi.Selection sel, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sel, cancel);
			Invoker.Method(this, "WindowBeforeDoubleClick", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}