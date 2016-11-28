using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.PublisherApi
{
	///<summary>
	/// DispatchInterface IApplicationEvents 
	/// SupportByVersion Publisher, 14,15,16
	///</summary>
	[SupportByVersionAttribute("Publisher", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class IApplicationEvents : COMObject
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
                    _type = typeof(IApplicationEvents);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IApplicationEvents(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wn">NetOffice.PublisherApi.Window Wn</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void WindowActivate(NetOffice.PublisherApi.Window wn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wn);
			Invoker.Method(this, "WindowActivate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wn">NetOffice.PublisherApi.Window Wn</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void WindowDeactivate(NetOffice.PublisherApi.Window wn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wn);
			Invoker.Method(this, "WindowDeactivate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="vw">NetOffice.PublisherApi.View Vw</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void WindowPageChange(NetOffice.PublisherApi.View vw)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(vw);
			Invoker.Method(this, "WindowPageChange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void Quit()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Quit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document Doc</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void NewDocument(NetOffice.PublisherApi._Document doc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc);
			Invoker.Method(this, "NewDocument", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document Doc</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void DocumentOpen(NetOffice.PublisherApi._Document doc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc);
			Invoker.Method(this, "DocumentOpen", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document Doc</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void DocumentBeforeClose(NetOffice.PublisherApi._Document doc, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, cancel);
			Invoker.Method(this, "DocumentBeforeClose", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document Doc</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void MailMergeAfterMerge(NetOffice.PublisherApi._Document doc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc);
			Invoker.Method(this, "MailMergeAfterMerge", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document Doc</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void MailMergeAfterRecordMerge(NetOffice.PublisherApi._Document doc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc);
			Invoker.Method(this, "MailMergeAfterRecordMerge", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document Doc</param>
		/// <param name="startRecord">Int32 StartRecord</param>
		/// <param name="endRecord">Int32 EndRecord</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void MailMergeBeforeMerge(NetOffice.PublisherApi._Document doc, Int32 startRecord, Int32 endRecord, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, startRecord, endRecord, cancel);
			Invoker.Method(this, "MailMergeBeforeMerge", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document Doc</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void MailMergeBeforeRecordMerge(NetOffice.PublisherApi._Document doc, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, cancel);
			Invoker.Method(this, "MailMergeBeforeRecordMerge", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document Doc</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void MailMergeDataSourceLoad(NetOffice.PublisherApi._Document doc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc);
			Invoker.Method(this, "MailMergeDataSourceLoad", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document Doc</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void MailMergeWizardSendToCustom(NetOffice.PublisherApi._Document doc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc);
			Invoker.Method(this, "MailMergeWizardSendToCustom", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document Doc</param>
		/// <param name="fromState">Int32 FromState</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void MailMergeWizardStateChange(NetOffice.PublisherApi._Document doc, Int32 fromState)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, fromState);
			Invoker.Method(this, "MailMergeWizardStateChange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document Doc</param>
		/// <param name="handled">bool Handled</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void MailMergeDataSourceValidate(NetOffice.PublisherApi._Document doc, bool handled)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, handled);
			Invoker.Method(this, "MailMergeDataSourceValidate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document Doc</param>
		/// <param name="okToInsert">bool OkToInsert</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void MailMergeInsertBarcode(NetOffice.PublisherApi._Document doc, bool okToInsert)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, okToInsert);
			Invoker.Method(this, "MailMergeInsertBarcode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document Doc</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void MailMergeRecipientListClose(NetOffice.PublisherApi._Document doc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc);
			Invoker.Method(this, "MailMergeRecipientListClose", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document Doc</param>
		/// <param name="bstrString">string bstrString</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void MailMergeGenerateBarcode(NetOffice.PublisherApi._Document doc, string bstrString)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, bstrString);
			Invoker.Method(this, "MailMergeGenerateBarcode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document Doc</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void MailMergeWizardFollowUpCustom(NetOffice.PublisherApi._Document doc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc);
			Invoker.Method(this, "MailMergeWizardFollowUpCustom", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document Doc</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void BeforePrint(NetOffice.PublisherApi._Document doc, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, cancel);
			Invoker.Method(this, "BeforePrint", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document Doc</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void AfterPrint(NetOffice.PublisherApi._Document doc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc);
			Invoker.Method(this, "AfterPrint", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ShowCatalogUI()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ShowCatalogUI", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void HideCatalogUI()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "HideCatalogUI", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}