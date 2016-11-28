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
	/// DispatchInterface IApplicationEvents4 
	/// SupportByVersion Word, 11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Word", 11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class IApplicationEvents4 : COMObject
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
                    _type = typeof(IApplicationEvents4);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IApplicationEvents4(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents4(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents4(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents4(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents4(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents4() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents4(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void Startup()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Startup", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void Quit()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Quit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void DocumentChange()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DocumentChange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void DocumentOpen(NetOffice.WordApi.Document doc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc);
			Invoker.Method(this, "DocumentOpen", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void DocumentBeforeClose(NetOffice.WordApi.Document doc, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, cancel);
			Invoker.Method(this, "DocumentBeforeClose", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void DocumentBeforePrint(NetOffice.WordApi.Document doc, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, cancel);
			Invoker.Method(this, "DocumentBeforePrint", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		/// <param name="saveAsUI">bool SaveAsUI</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void DocumentBeforeSave(NetOffice.WordApi.Document doc, bool saveAsUI, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, saveAsUI, cancel);
			Invoker.Method(this, "DocumentBeforeSave", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void NewDocument(NetOffice.WordApi.Document doc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc);
			Invoker.Method(this, "NewDocument", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		/// <param name="wn">NetOffice.WordApi.Window Wn</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void WindowActivate(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, wn);
			Invoker.Method(this, "WindowActivate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		/// <param name="wn">NetOffice.WordApi.Window Wn</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void WindowDeactivate(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, wn);
			Invoker.Method(this, "WindowDeactivate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection Sel</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void WindowSelectionChange(NetOffice.WordApi.Selection sel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sel);
			Invoker.Method(this, "WindowSelectionChange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection Sel</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void WindowBeforeRightClick(NetOffice.WordApi.Selection sel, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sel, cancel);
			Invoker.Method(this, "WindowBeforeRightClick", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection Sel</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void WindowBeforeDoubleClick(NetOffice.WordApi.Selection sel, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sel, cancel);
			Invoker.Method(this, "WindowBeforeDoubleClick", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void EPostagePropertyDialog(NetOffice.WordApi.Document doc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc);
			Invoker.Method(this, "EPostagePropertyDialog", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void EPostageInsert(NetOffice.WordApi.Document doc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc);
			Invoker.Method(this, "EPostageInsert", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		/// <param name="docResult">NetOffice.WordApi.Document DocResult</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void MailMergeAfterMerge(NetOffice.WordApi.Document doc, NetOffice.WordApi.Document docResult)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, docResult);
			Invoker.Method(this, "MailMergeAfterMerge", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void MailMergeAfterRecordMerge(NetOffice.WordApi.Document doc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc);
			Invoker.Method(this, "MailMergeAfterRecordMerge", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		/// <param name="startRecord">Int32 StartRecord</param>
		/// <param name="endRecord">Int32 EndRecord</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void MailMergeBeforeMerge(NetOffice.WordApi.Document doc, Int32 startRecord, Int32 endRecord, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, startRecord, endRecord, cancel);
			Invoker.Method(this, "MailMergeBeforeMerge", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void MailMergeBeforeRecordMerge(NetOffice.WordApi.Document doc, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, cancel);
			Invoker.Method(this, "MailMergeBeforeRecordMerge", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void MailMergeDataSourceLoad(NetOffice.WordApi.Document doc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc);
			Invoker.Method(this, "MailMergeDataSourceLoad", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		/// <param name="handled">bool Handled</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void MailMergeDataSourceValidate(NetOffice.WordApi.Document doc, bool handled)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, handled);
			Invoker.Method(this, "MailMergeDataSourceValidate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void MailMergeWizardSendToCustom(NetOffice.WordApi.Document doc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc);
			Invoker.Method(this, "MailMergeWizardSendToCustom", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		/// <param name="fromState">Int32 FromState</param>
		/// <param name="toState">Int32 ToState</param>
		/// <param name="handled">bool Handled</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void MailMergeWizardStateChange(NetOffice.WordApi.Document doc, Int32 fromState, Int32 toState, bool handled)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, fromState, toState, handled);
			Invoker.Method(this, "MailMergeWizardStateChange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		/// <param name="wn">NetOffice.WordApi.Window Wn</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void WindowSize(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, wn);
			Invoker.Method(this, "WindowSize", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection Sel</param>
		/// <param name="oldXMLNode">NetOffice.WordApi.XMLNode OldXMLNode</param>
		/// <param name="newXMLNode">NetOffice.WordApi.XMLNode NewXMLNode</param>
		/// <param name="reason">Int32 Reason</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void XMLSelectionChange(NetOffice.WordApi.Selection sel, NetOffice.WordApi.XMLNode oldXMLNode, NetOffice.WordApi.XMLNode newXMLNode, Int32 reason)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sel, oldXMLNode, newXMLNode, reason);
			Invoker.Method(this, "XMLSelectionChange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xMLNode">NetOffice.WordApi.XMLNode XMLNode</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void XMLValidationError(NetOffice.WordApi.XMLNode xMLNode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xMLNode);
			Invoker.Method(this, "XMLValidationError", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		/// <param name="syncEventType">NetOffice.OfficeApi.Enums.MsoSyncEventType SyncEventType</param>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void DocumentSync(NetOffice.WordApi.Document doc, NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, syncEventType);
			Invoker.Method(this, "DocumentSync", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
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
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void EPostageInsertEx(NetOffice.WordApi.Document doc, Int32 cpDeliveryAddrStart, Int32 cpDeliveryAddrEnd, Int32 cpReturnAddrStart, Int32 cpReturnAddrEnd, Int32 xaWidth, Int32 yaHeight, string bstrPrinterName, string bstrPaperFeed, bool fPrint, bool fCancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, cpDeliveryAddrStart, cpDeliveryAddrEnd, cpReturnAddrStart, cpReturnAddrEnd, xaWidth, yaHeight, bstrPrinterName, bstrPaperFeed, fPrint, fCancel);
			Invoker.Method(this, "EPostageInsertEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document Doc</param>
		/// <param name="handled">bool Handled</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public Int32 MailMergeDataSourceValidate2(NetOffice.WordApi.Document doc, bool handled)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(doc, handled);
			object returnItem = Invoker.MethodReturn(this, "MailMergeDataSourceValidate2", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pvWindow">NetOffice.WordApi.ProtectedViewWindow PvWindow</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public Int32 ProtectedViewWindowOpen(NetOffice.WordApi.ProtectedViewWindow pvWindow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvWindow);
			object returnItem = Invoker.MethodReturn(this, "ProtectedViewWindowOpen", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pvWindow">NetOffice.WordApi.ProtectedViewWindow PvWindow</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public Int32 ProtectedViewWindowBeforeEdit(NetOffice.WordApi.ProtectedViewWindow pvWindow, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvWindow, cancel);
			object returnItem = Invoker.MethodReturn(this, "ProtectedViewWindowBeforeEdit", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pvWindow">NetOffice.WordApi.ProtectedViewWindow PvWindow</param>
		/// <param name="closeReason">Int32 CloseReason</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public Int32 ProtectedViewWindowBeforeClose(NetOffice.WordApi.ProtectedViewWindow pvWindow, Int32 closeReason, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvWindow, closeReason, cancel);
			object returnItem = Invoker.MethodReturn(this, "ProtectedViewWindowBeforeClose", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pvWindow">NetOffice.WordApi.ProtectedViewWindow PvWindow</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public Int32 ProtectedViewWindowSize(NetOffice.WordApi.ProtectedViewWindow pvWindow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvWindow);
			object returnItem = Invoker.MethodReturn(this, "ProtectedViewWindowSize", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pvWindow">NetOffice.WordApi.ProtectedViewWindow PvWindow</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public Int32 ProtectedViewWindowActivate(NetOffice.WordApi.ProtectedViewWindow pvWindow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvWindow);
			object returnItem = Invoker.MethodReturn(this, "ProtectedViewWindowActivate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pvWindow">NetOffice.WordApi.ProtectedViewWindow PvWindow</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public Int32 ProtectedViewWindowDeactivate(NetOffice.WordApi.ProtectedViewWindow pvWindow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvWindow);
			object returnItem = Invoker.MethodReturn(this, "ProtectedViewWindowDeactivate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}