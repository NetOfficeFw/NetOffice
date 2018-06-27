using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PublisherApi;

namespace NetOffice.PublisherApi.Behind
{
	/// <summary>
	/// DispatchInterface IApplicationEvents 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IApplicationEvents : COMObject, NetOffice.PublisherApi.IApplicationEvents
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
                    _contractType = typeof(NetOffice.PublisherApi.IApplicationEvents);
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
                    _type = typeof(IApplicationEvents);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IApplicationEvents() : base()
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
		public virtual void WindowActivate(NetOffice.PublisherApi.Window wn)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WindowActivate", wn);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wn">NetOffice.PublisherApi.Window wn</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void WindowDeactivate(NetOffice.PublisherApi.Window wn)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WindowDeactivate", wn);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="vw">NetOffice.PublisherApi.View vw</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void WindowPageChange(NetOffice.PublisherApi.View vw)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WindowPageChange", vw);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Quit()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Quit");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void NewDocument(NetOffice.PublisherApi._Document doc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "NewDocument", doc);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void DocumentOpen(NetOffice.PublisherApi._Document doc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DocumentOpen", doc);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void DocumentBeforeClose(NetOffice.PublisherApi._Document doc, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DocumentBeforeClose", doc, cancel);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void MailMergeAfterMerge(NetOffice.PublisherApi._Document doc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeAfterMerge", doc);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void MailMergeAfterRecordMerge(NetOffice.PublisherApi._Document doc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeAfterRecordMerge", doc);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="startRecord">Int32 startRecord</param>
		/// <param name="endRecord">Int32 endRecord</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void MailMergeBeforeMerge(NetOffice.PublisherApi._Document doc, Int32 startRecord, Int32 endRecord, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeBeforeMerge", doc, startRecord, endRecord, cancel);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void MailMergeBeforeRecordMerge(NetOffice.PublisherApi._Document doc, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeBeforeRecordMerge", doc, cancel);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void MailMergeDataSourceLoad(NetOffice.PublisherApi._Document doc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeDataSourceLoad", doc);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void MailMergeWizardSendToCustom(NetOffice.PublisherApi._Document doc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeWizardSendToCustom", doc);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="fromState">Int32 fromState</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void MailMergeWizardStateChange(NetOffice.PublisherApi._Document doc, Int32 fromState)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeWizardStateChange", doc, fromState);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="handled">bool handled</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void MailMergeDataSourceValidate(NetOffice.PublisherApi._Document doc, bool handled)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeDataSourceValidate", doc, handled);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="okToInsert">bool okToInsert</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void MailMergeInsertBarcode(NetOffice.PublisherApi._Document doc, bool okToInsert)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeInsertBarcode", doc, okToInsert);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void MailMergeRecipientListClose(NetOffice.PublisherApi._Document doc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeRecipientListClose", doc);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="bstrString">string bstrString</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void MailMergeGenerateBarcode(NetOffice.PublisherApi._Document doc, string bstrString)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeGenerateBarcode", doc, bstrString);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void MailMergeWizardFollowUpCustom(NetOffice.PublisherApi._Document doc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeWizardFollowUpCustom", doc);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void BeforePrint(NetOffice.PublisherApi._Document doc, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BeforePrint", doc, cancel);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.PublisherApi._Document doc</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void AfterPrint(NetOffice.PublisherApi._Document doc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AfterPrint", doc);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ShowCatalogUI()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowCatalogUI");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void HideCatalogUI()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "HideCatalogUI");
		}

		#endregion

		#pragma warning restore
	}
}

