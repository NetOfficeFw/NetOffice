using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface IApplicationEvents4 
	/// SupportByVersion Word, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Word", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IApplicationEvents4 : COMObject, NetOffice.WordApi.IApplicationEvents4
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
                    _contractType = typeof(NetOffice.WordApi.IApplicationEvents4);
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
                    _type = typeof(IApplicationEvents4);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IApplicationEvents4() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void Startup()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Startup");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void Quit()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Quit");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void DocumentChange()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DocumentChange");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void DocumentOpen(NetOffice.WordApi.Document doc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DocumentOpen", doc);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void DocumentBeforeClose(NetOffice.WordApi.Document doc, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DocumentBeforeClose", doc, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void DocumentBeforePrint(NetOffice.WordApi.Document doc, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DocumentBeforePrint", doc, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="saveAsUI">bool saveAsUI</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void DocumentBeforeSave(NetOffice.WordApi.Document doc, bool saveAsUI, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DocumentBeforeSave", doc, saveAsUI, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void NewDocument(NetOffice.WordApi.Document doc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "NewDocument", doc);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="wn">NetOffice.WordApi.Window wn</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void WindowActivate(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WindowActivate", doc, wn);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="wn">NetOffice.WordApi.Window wn</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void WindowDeactivate(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WindowDeactivate", doc, wn);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void WindowSelectionChange(NetOffice.WordApi.Selection sel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WindowSelectionChange", sel);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void WindowBeforeRightClick(NetOffice.WordApi.Selection sel, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WindowBeforeRightClick", sel, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void WindowBeforeDoubleClick(NetOffice.WordApi.Selection sel, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WindowBeforeDoubleClick", sel, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void EPostagePropertyDialog(NetOffice.WordApi.Document doc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EPostagePropertyDialog", doc);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void EPostageInsert(NetOffice.WordApi.Document doc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EPostageInsert", doc);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="docResult">NetOffice.WordApi.Document docResult</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void MailMergeAfterMerge(NetOffice.WordApi.Document doc, NetOffice.WordApi.Document docResult)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeAfterMerge", doc, docResult);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void MailMergeAfterRecordMerge(NetOffice.WordApi.Document doc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeAfterRecordMerge", doc);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="startRecord">Int32 startRecord</param>
		/// <param name="endRecord">Int32 endRecord</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void MailMergeBeforeMerge(NetOffice.WordApi.Document doc, Int32 startRecord, Int32 endRecord, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeBeforeMerge", doc, startRecord, endRecord, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void MailMergeBeforeRecordMerge(NetOffice.WordApi.Document doc, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeBeforeRecordMerge", doc, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void MailMergeDataSourceLoad(NetOffice.WordApi.Document doc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeDataSourceLoad", doc);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="handled">bool handled</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void MailMergeDataSourceValidate(NetOffice.WordApi.Document doc, bool handled)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeDataSourceValidate", doc, handled);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void MailMergeWizardSendToCustom(NetOffice.WordApi.Document doc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeWizardSendToCustom", doc);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="fromState">Int32 fromState</param>
		/// <param name="toState">Int32 toState</param>
		/// <param name="handled">bool handled</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void MailMergeWizardStateChange(NetOffice.WordApi.Document doc, Int32 fromState, Int32 toState, bool handled)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MailMergeWizardStateChange", doc, fromState, toState, handled);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="wn">NetOffice.WordApi.Window wn</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void WindowSize(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WindowSize", doc, wn);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		/// <param name="oldXMLNode">NetOffice.WordApi.XMLNode oldXMLNode</param>
		/// <param name="newXMLNode">NetOffice.WordApi.XMLNode newXMLNode</param>
		/// <param name="reason">Int32 reason</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void XMLSelectionChange(NetOffice.WordApi.Selection sel, NetOffice.WordApi.XMLNode oldXMLNode, NetOffice.WordApi.XMLNode newXMLNode, Int32 reason)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "XMLSelectionChange", sel, oldXMLNode, newXMLNode, reason);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xMLNode">NetOffice.WordApi.XMLNode xMLNode</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void XMLValidationError(NetOffice.WordApi.XMLNode xMLNode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "XMLValidationError", xMLNode);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="syncEventType">NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void DocumentSync(NetOffice.WordApi.Document doc, NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DocumentSync", doc, syncEventType);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="cpDeliveryAddrStart">Int32 cpDeliveryAddrStart</param>
		/// <param name="cpDeliveryAddrEnd">Int32 cpDeliveryAddrEnd</param>
		/// <param name="cpReturnAddrStart">Int32 cpReturnAddrStart</param>
		/// <param name="cpReturnAddrEnd">Int32 cpReturnAddrEnd</param>
		/// <param name="xaWidth">Int32 xaWidth</param>
		/// <param name="yaHeight">Int32 yaHeight</param>
		/// <param name="bstrPrinterName">string bstrPrinterName</param>
		/// <param name="bstrPaperFeed">string bstrPaperFeed</param>
		/// <param name="fPrint">bool fPrint</param>
		/// <param name="fCancel">bool fCancel</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void EPostageInsertEx(NetOffice.WordApi.Document doc, Int32 cpDeliveryAddrStart, Int32 cpDeliveryAddrEnd, Int32 cpReturnAddrStart, Int32 cpReturnAddrEnd, Int32 xaWidth, Int32 yaHeight, string bstrPrinterName, string bstrPaperFeed, bool fPrint, bool fCancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EPostageInsertEx", new object[]{ doc, cpDeliveryAddrStart, cpDeliveryAddrEnd, cpReturnAddrStart, cpReturnAddrEnd, xaWidth, yaHeight, bstrPrinterName, bstrPaperFeed, fPrint, fCancel });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="handled">bool handled</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual Int32 MailMergeDataSourceValidate2(NetOffice.WordApi.Document doc, bool handled)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MailMergeDataSourceValidate2", doc, handled);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="pvWindow">NetOffice.WordApi.ProtectedViewWindow pvWindow</param>
		[SupportByVersion("Word", 14,15,16)]
		public virtual Int32 ProtectedViewWindowOpen(NetOffice.WordApi.ProtectedViewWindow pvWindow)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ProtectedViewWindowOpen", pvWindow);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="pvWindow">NetOffice.WordApi.ProtectedViewWindow pvWindow</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 14,15,16)]
		public virtual Int32 ProtectedViewWindowBeforeEdit(NetOffice.WordApi.ProtectedViewWindow pvWindow, bool cancel)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ProtectedViewWindowBeforeEdit", pvWindow, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="pvWindow">NetOffice.WordApi.ProtectedViewWindow pvWindow</param>
		/// <param name="closeReason">Int32 closeReason</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 14,15,16)]
		public virtual Int32 ProtectedViewWindowBeforeClose(NetOffice.WordApi.ProtectedViewWindow pvWindow, Int32 closeReason, bool cancel)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ProtectedViewWindowBeforeClose", pvWindow, closeReason, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="pvWindow">NetOffice.WordApi.ProtectedViewWindow pvWindow</param>
		[SupportByVersion("Word", 14,15,16)]
		public virtual Int32 ProtectedViewWindowSize(NetOffice.WordApi.ProtectedViewWindow pvWindow)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ProtectedViewWindowSize", pvWindow);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="pvWindow">NetOffice.WordApi.ProtectedViewWindow pvWindow</param>
		[SupportByVersion("Word", 14,15,16)]
		public virtual Int32 ProtectedViewWindowActivate(NetOffice.WordApi.ProtectedViewWindow pvWindow)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ProtectedViewWindowActivate", pvWindow);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="pvWindow">NetOffice.WordApi.ProtectedViewWindow pvWindow</param>
		[SupportByVersion("Word", 14,15,16)]
		public virtual Int32 ProtectedViewWindowDeactivate(NetOffice.WordApi.ProtectedViewWindow pvWindow)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ProtectedViewWindowDeactivate", pvWindow);
		}

		#endregion

		#pragma warning restore
	}
}

