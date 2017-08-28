using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface IApplicationEvents 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IApplicationEvents : COMObject
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
                    _type = typeof(IApplicationEvents);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IApplicationEvents(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		/// </summary>
		/// <param name="wn">NetOffice.PublisherApi.Window wn</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void WindowActivate(NetOffice.PublisherApi.Window wn)
		{
			 Factory.ExecuteMethod(this, "WindowActivate", wn);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wn">NetOffice.PublisherApi.Window wn</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void WindowDeactivate(NetOffice.PublisherApi.Window wn)
		{
			 Factory.ExecuteMethod(this, "WindowDeactivate", wn);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="vw">NetOffice.PublisherApi.View vw</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void WindowPageChange(NetOffice.PublisherApi.View vw)
		{
			 Factory.ExecuteMethod(this, "WindowPageChange", vw);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Quit()
		{
			 Factory.ExecuteMethod(this, "Quit");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void NewDocument(NetOffice.PublisherApi._Document doc)
		{
			 Factory.ExecuteMethod(this, "NewDocument", doc);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void DocumentOpen(NetOffice.PublisherApi._Document doc)
		{
			 Factory.ExecuteMethod(this, "DocumentOpen", doc);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void DocumentBeforeClose(NetOffice.PublisherApi._Document doc, bool cancel)
		{
			 Factory.ExecuteMethod(this, "DocumentBeforeClose", doc, cancel);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void MailMergeAfterMerge(NetOffice.PublisherApi._Document doc)
		{
			 Factory.ExecuteMethod(this, "MailMergeAfterMerge", doc);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void MailMergeAfterRecordMerge(NetOffice.PublisherApi._Document doc)
		{
			 Factory.ExecuteMethod(this, "MailMergeAfterRecordMerge", doc);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="startRecord">Int32 startRecord</param>
		/// <param name="endRecord">Int32 endRecord</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void MailMergeBeforeMerge(NetOffice.PublisherApi._Document doc, Int32 startRecord, Int32 endRecord, bool cancel)
		{
			 Factory.ExecuteMethod(this, "MailMergeBeforeMerge", doc, startRecord, endRecord, cancel);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void MailMergeBeforeRecordMerge(NetOffice.PublisherApi._Document doc, bool cancel)
		{
			 Factory.ExecuteMethod(this, "MailMergeBeforeRecordMerge", doc, cancel);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void MailMergeDataSourceLoad(NetOffice.PublisherApi._Document doc)
		{
			 Factory.ExecuteMethod(this, "MailMergeDataSourceLoad", doc);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		public void MailMergeWizardSendToCustom(NetOffice.PublisherApi._Document doc)
		{
			 Factory.ExecuteMethod(this, "MailMergeWizardSendToCustom", doc);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="fromState">Int32 fromState</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void MailMergeWizardStateChange(NetOffice.PublisherApi._Document doc, Int32 fromState)
		{
			 Factory.ExecuteMethod(this, "MailMergeWizardStateChange", doc, fromState);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="handled">bool handled</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void MailMergeDataSourceValidate(NetOffice.PublisherApi._Document doc, bool handled)
		{
			 Factory.ExecuteMethod(this, "MailMergeDataSourceValidate", doc, handled);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="okToInsert">bool okToInsert</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void MailMergeInsertBarcode(NetOffice.PublisherApi._Document doc, bool okToInsert)
		{
			 Factory.ExecuteMethod(this, "MailMergeInsertBarcode", doc, okToInsert);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void MailMergeRecipientListClose(NetOffice.PublisherApi._Document doc)
		{
			 Factory.ExecuteMethod(this, "MailMergeRecipientListClose", doc);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="bstrString">string bstrString</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void MailMergeGenerateBarcode(NetOffice.PublisherApi._Document doc, string bstrString)
		{
			 Factory.ExecuteMethod(this, "MailMergeGenerateBarcode", doc, bstrString);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void MailMergeWizardFollowUpCustom(NetOffice.PublisherApi._Document doc)
		{
			 Factory.ExecuteMethod(this, "MailMergeWizardFollowUpCustom", doc);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void BeforePrint(NetOffice.PublisherApi._Document doc, bool cancel)
		{
			 Factory.ExecuteMethod(this, "BeforePrint", doc, cancel);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void AfterPrint(NetOffice.PublisherApi._Document doc)
		{
			 Factory.ExecuteMethod(this, "AfterPrint", doc);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void ShowCatalogUI()
		{
			 Factory.ExecuteMethod(this, "ShowCatalogUI");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void HideCatalogUI()
		{
			 Factory.ExecuteMethod(this, "HideCatalogUI");
		}

		#endregion

		#pragma warning restore
	}
}
