using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface IApplicationEvents3 
	/// SupportByVersion Word, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Word", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IApplicationEvents3 : COMObject
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
                    _type = typeof(IApplicationEvents3);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IApplicationEvents3(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IApplicationEvents3(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents3(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents3(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents3(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents3(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents3() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IApplicationEvents3(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Startup()
		{
			 Factory.ExecuteMethod(this, "Startup");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Quit()
		{
			 Factory.ExecuteMethod(this, "Quit");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void DocumentChange()
		{
			 Factory.ExecuteMethod(this, "DocumentChange");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void DocumentOpen(NetOffice.WordApi.Document doc)
		{
			 Factory.ExecuteMethod(this, "DocumentOpen", doc);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void DocumentBeforeClose(NetOffice.WordApi.Document doc, bool cancel)
		{
			 Factory.ExecuteMethod(this, "DocumentBeforeClose", doc, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void DocumentBeforePrint(NetOffice.WordApi.Document doc, bool cancel)
		{
			 Factory.ExecuteMethod(this, "DocumentBeforePrint", doc, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="saveAsUI">bool saveAsUI</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void DocumentBeforeSave(NetOffice.WordApi.Document doc, bool saveAsUI, bool cancel)
		{
			 Factory.ExecuteMethod(this, "DocumentBeforeSave", doc, saveAsUI, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void NewDocument(NetOffice.WordApi.Document doc)
		{
			 Factory.ExecuteMethod(this, "NewDocument", doc);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="wn">NetOffice.WordApi.Window wn</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void WindowActivate(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn)
		{
			 Factory.ExecuteMethod(this, "WindowActivate", doc, wn);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="wn">NetOffice.WordApi.Window wn</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void WindowDeactivate(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn)
		{
			 Factory.ExecuteMethod(this, "WindowDeactivate", doc, wn);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void WindowSelectionChange(NetOffice.WordApi.Selection sel)
		{
			 Factory.ExecuteMethod(this, "WindowSelectionChange", sel);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void WindowBeforeRightClick(NetOffice.WordApi.Selection sel, bool cancel)
		{
			 Factory.ExecuteMethod(this, "WindowBeforeRightClick", sel, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void WindowBeforeDoubleClick(NetOffice.WordApi.Selection sel, bool cancel)
		{
			 Factory.ExecuteMethod(this, "WindowBeforeDoubleClick", sel, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void EPostagePropertyDialog(NetOffice.WordApi.Document doc)
		{
			 Factory.ExecuteMethod(this, "EPostagePropertyDialog", doc);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void EPostageInsert(NetOffice.WordApi.Document doc)
		{
			 Factory.ExecuteMethod(this, "EPostageInsert", doc);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="docResult">NetOffice.WordApi.Document docResult</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void MailMergeAfterMerge(NetOffice.WordApi.Document doc, NetOffice.WordApi.Document docResult)
		{
			 Factory.ExecuteMethod(this, "MailMergeAfterMerge", doc, docResult);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void MailMergeAfterRecordMerge(NetOffice.WordApi.Document doc)
		{
			 Factory.ExecuteMethod(this, "MailMergeAfterRecordMerge", doc);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="startRecord">Int32 startRecord</param>
		/// <param name="endRecord">Int32 endRecord</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void MailMergeBeforeMerge(NetOffice.WordApi.Document doc, Int32 startRecord, Int32 endRecord, bool cancel)
		{
			 Factory.ExecuteMethod(this, "MailMergeBeforeMerge", doc, startRecord, endRecord, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void MailMergeBeforeRecordMerge(NetOffice.WordApi.Document doc, bool cancel)
		{
			 Factory.ExecuteMethod(this, "MailMergeBeforeRecordMerge", doc, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void MailMergeDataSourceLoad(NetOffice.WordApi.Document doc)
		{
			 Factory.ExecuteMethod(this, "MailMergeDataSourceLoad", doc);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="handled">bool handled</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void MailMergeDataSourceValidate(NetOffice.WordApi.Document doc, bool handled)
		{
			 Factory.ExecuteMethod(this, "MailMergeDataSourceValidate", doc, handled);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void MailMergeWizardSendToCustom(NetOffice.WordApi.Document doc)
		{
			 Factory.ExecuteMethod(this, "MailMergeWizardSendToCustom", doc);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="fromState">Int32 fromState</param>
		/// <param name="toState">Int32 toState</param>
		/// <param name="handled">bool handled</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void MailMergeWizardStateChange(NetOffice.WordApi.Document doc, Int32 fromState, Int32 toState, bool handled)
		{
			 Factory.ExecuteMethod(this, "MailMergeWizardStateChange", doc, fromState, toState, handled);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="wn">NetOffice.WordApi.Window wn</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void WindowSize(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn)
		{
			 Factory.ExecuteMethod(this, "WindowSize", doc, wn);
		}

		#endregion

		#pragma warning restore
	}
}
