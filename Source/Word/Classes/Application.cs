﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.WordApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Application_StartupEventHandler();
	public delegate void Application_QuitEventHandler();
	public delegate void Application_DocumentChangeEventHandler();
	public delegate void Application_DocumentOpenEventHandler(NetOffice.WordApi.Document doc);
	public delegate void Application_DocumentBeforeCloseEventHandler(NetOffice.WordApi.Document doc, ref bool cancel);
	public delegate void Application_DocumentBeforePrintEventHandler(NetOffice.WordApi.Document doc, ref bool cancel);
	public delegate void Application_DocumentBeforeSaveEventHandler(NetOffice.WordApi.Document doc, ref bool saveAsUI, ref bool cancel);
	public delegate void Application_NewDocumentEventHandler(NetOffice.WordApi.Document doc);
	public delegate void Application_WindowActivateEventHandler(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn);
	public delegate void Application_WindowDeactivateEventHandler(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn);
	public delegate void Application_WindowSelectionChangeEventHandler(NetOffice.WordApi.Selection sel);
	public delegate void Application_WindowBeforeRightClickEventHandler(NetOffice.WordApi.Selection sel, ref bool cancel);
	public delegate void Application_WindowBeforeDoubleClickEventHandler(NetOffice.WordApi.Selection sel, ref bool cancel);
	public delegate void Application_EPostagePropertyDialogEventHandler(NetOffice.WordApi.Document doc);
	public delegate void Application_EPostageInsertEventHandler(NetOffice.WordApi.Document doc);
	public delegate void Application_MailMergeAfterMergeEventHandler(NetOffice.WordApi.Document doc, NetOffice.WordApi.Document docResult);
	public delegate void Application_MailMergeAfterRecordMergeEventHandler(NetOffice.WordApi.Document doc);
	public delegate void Application_MailMergeBeforeMergeEventHandler(NetOffice.WordApi.Document doc, Int32 startRecord, Int32 endRecord, ref bool cancel);
	public delegate void Application_MailMergeBeforeRecordMergeEventHandler(NetOffice.WordApi.Document doc, ref bool cancel);
	public delegate void Application_MailMergeDataSourceLoadEventHandler(NetOffice.WordApi.Document doc);
	public delegate void Application_MailMergeDataSourceValidateEventHandler(NetOffice.WordApi.Document doc, ref bool handled);
	public delegate void Application_MailMergeWizardSendToCustomEventHandler(NetOffice.WordApi.Document doc);
	public delegate void Application_MailMergeWizardStateChangeEventHandler(NetOffice.WordApi.Document doc, ref Int32 fromState, ref Int32 toState, ref bool handled);
	public delegate void Application_WindowSizeEventHandler(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn);
	public delegate void Application_XMLSelectionChangeEventHandler(NetOffice.WordApi.Selection sel, NetOffice.WordApi.XMLNode oldXMLNode, NetOffice.WordApi.XMLNode newXMLNode, ref Int32 reason);
	public delegate void Application_XMLValidationErrorEventHandler(NetOffice.WordApi.XMLNode xmlNode);
	public delegate void Application_DocumentSyncEventHandler(NetOffice.WordApi.Document doc, NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType);
	public delegate void Application_EPostageInsertExEventHandler(NetOffice.WordApi.Document doc, Int32 cpDeliveryAddrStart, Int32 cpDeliveryAddrEnd, Int32 cpReturnAddrStart, Int32 cpReturnAddrEnd, Int32 xaWidth, Int32 yaHeight, string bstrPrinterName, string bstrPaperFeed, bool fPrint, ref bool fCancel);
	public delegate void Application_MailMergeDataSourceValidate2EventHandler(NetOffice.WordApi.Document doc, ref bool handled);
	public delegate void Application_ProtectedViewWindowOpenEventHandler(NetOffice.WordApi.ProtectedViewWindow pvWindow);
	public delegate void Application_ProtectedViewWindowBeforeEditEventHandler(NetOffice.WordApi.ProtectedViewWindow pvWindow, ref bool cancel);
	public delegate void Application_ProtectedViewWindowBeforeCloseEventHandler(NetOffice.WordApi.ProtectedViewWindow pvWindow, Int32 closeReason, ref bool cancel);
	public delegate void Application_ProtectedViewWindowSizeEventHandler(NetOffice.WordApi.ProtectedViewWindow pvWindow);
	public delegate void Application_ProtectedViewWindowActivateEventHandler(NetOffice.WordApi.ProtectedViewWindow pvWindow);
	public delegate void Application_ProtectedViewWindowDeactivateEventHandler(NetOffice.WordApi.ProtectedViewWindow pvWindow);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Application 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838565.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass), ComProgId("Word.Application"), ModuleProvider(typeof(GlobalHelperModules.GlobalModule))]
	[EventSink(typeof(Events.ApplicationEvents2_SinkHelper), typeof(Events.ApplicationEvents3_SinkHelper), typeof(Events.ApplicationEvents4_SinkHelper))]
    [ComEventInterface(typeof(Events.ApplicationEvents2), typeof(Events.ApplicationEvents3), typeof(Events.ApplicationEvents4))]
    public class Application : _Application, ICloneable<Application>, IEventBinding
	{
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events.ApplicationEvents2_SinkHelper _applicationEvents2_SinkHelper;
        private Events.ApplicationEvents3_SinkHelper _applicationEvents3_SinkHelper;
		private Events.ApplicationEvents4_SinkHelper _applicationEvents4_SinkHelper;
	
		#endregion

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
        
        /// <summary>
        /// Type Cache
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Application);
                return _type;
            }
        }
        
        #endregion
        		
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Application(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
			_callQuitInDispose = true;
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Application(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			_callQuitInDispose = true;
			GlobalHelperModules.GlobalModule.Instance = this;
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Application(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			_callQuitInDispose = true;
			GlobalHelperModules.GlobalModule.Instance = this;
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Application(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			_callQuitInDispose = true;
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Application(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			_callQuitInDispose = true;
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Application(ICOMObject replacedObject) : base(replacedObject)
		{
			_callQuitInDispose = true;
		}
		
		/// <summary>
        /// Creates a new instance of Application 
        /// </summary>		
		public Application() : this(false)
		{

		}

        /// <summary>
        /// Creates a new instance of Application 
        /// <param name="enableProxyService">try to get a running application first before create a new application</param>
        /// </summary>		
        public Application(bool enableProxyService = false) : base()
        {
            if (enableProxyService)
            {
                Factory = Core.Default;
                object proxy = Running.ProxyService.GetActiveInstance("Word", "Application", false);
                if (null != proxy)
                {
                    CreateFromProxy(proxy, true);
                    FromProxyService = true;
                }
                else
                {
                    CreateFromProgId("Word.Application", true);
                }
            }
            else
            {
                CreateFromProgId("Word.Application", true);
            }

            OnCreate();
            _callQuitInDispose = true;
            GlobalHelperModules.GlobalModule.Instance = this;
        }

        /// <summary>
        /// Creates a new instance of Application
        /// </summary>
        ///<param name="progId">registered ProgID</param>
        public Application(string progId):base(progId)
		{
			_callQuitInDispose = true;
			GlobalHelperModules.GlobalModule.Instance = this;
		}
		
        /// <summary>
		/// NetOffice method: dispose instance and all child instances
		/// </summary>
		/// <param name="disposeEventBinding">dispose event exported proxies with one or more event recipients</param>
		[Category("NetOffice"), CoreOverridden]
		public override void Dispose(bool disposeEventBinding)
		{
			if(this.Equals(GlobalHelperModules.GlobalModule.Instance))
				 GlobalHelperModules.GlobalModule.Instance = null;	
			base.Dispose(disposeEventBinding);
		}

		/// <summary>
		/// NetOffice method: dispose instance and all child instances
		/// </summary>
		[Category("NetOffice"), CoreOverridden]
		public override void Dispose()
		{
			if(this.Equals(GlobalHelperModules.GlobalModule.Instance))
				 GlobalHelperModules.GlobalModule.Instance = null;
			base.Dispose();
		}

        #endregion
        
        #region Properties

        /// <summary>
        /// Instance is created from an already running application
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public bool FromProxyService { get; private set; }

        #endregion

        #region Static CoClass Methods

        /// <summary>
        /// Returns all running Word.Application instances from the environment/system
        /// </summary>
        /// <returns>Word.Application sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances()
        {
            return Running.ProxyService.GetActiveInstances<Application>("Word", "Application");
        }

        /// <summary>
        /// Returns a running Word.Application instance from the environment/system
        /// </summary>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>Word.Application instance or null</returns>
        public static Application GetActiveInstance(bool throwExceptionIfNotFound = false)
        {
            return Running.ProxyService.GetActiveInstance<Application>("Word", "Application", throwExceptionIfNotFound);
        }

        #endregion

        #region Events

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        private event Application_StartupEventHandler _StartupEvent;

		/// <summary>
		/// SupportByVersion Word 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public event Application_StartupEventHandler StartupEvent
		{
			add
			{
				CreateEventBridge();
				_StartupEvent += value;
			}
			remove
			{
				_StartupEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_QuitEventHandler _QuitEvent;

		/// <summary>
		/// SupportByVersion Word 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194164.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public event Application_QuitEventHandler QuitEvent
		{
			add
			{
				CreateEventBridge();
				_QuitEvent += value;
			}
			remove
			{
				_QuitEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_DocumentChangeEventHandler _DocumentChangeEvent;

		/// <summary>
		/// SupportByVersion Word 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822189.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public event Application_DocumentChangeEventHandler DocumentChangeEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentChangeEvent += value;
			}
			remove
			{
				_DocumentChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_DocumentOpenEventHandler _DocumentOpenEvent;

		/// <summary>
		/// SupportByVersion Word 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192207.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public event Application_DocumentOpenEventHandler DocumentOpenEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentOpenEvent += value;
			}
			remove
			{
				_DocumentOpenEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_DocumentBeforeCloseEventHandler _DocumentBeforeCloseEvent;

		/// <summary>
		/// SupportByVersion Word 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834271.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public event Application_DocumentBeforeCloseEventHandler DocumentBeforeCloseEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentBeforeCloseEvent += value;
			}
			remove
			{
				_DocumentBeforeCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_DocumentBeforePrintEventHandler _DocumentBeforePrintEvent;

		/// <summary>
		/// SupportByVersion Word 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845163.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public event Application_DocumentBeforePrintEventHandler DocumentBeforePrintEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentBeforePrintEvent += value;
			}
			remove
			{
				_DocumentBeforePrintEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_DocumentBeforeSaveEventHandler _DocumentBeforeSaveEvent;

		/// <summary>
		/// SupportByVersion Word 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838299.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public event Application_DocumentBeforeSaveEventHandler DocumentBeforeSaveEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentBeforeSaveEvent += value;
			}
			remove
			{
				_DocumentBeforeSaveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_NewDocumentEventHandler _NewDocumentEvent;

		/// <summary>
		/// SupportByVersion Word 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836563.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public event Application_NewDocumentEventHandler NewDocumentEvent
		{
			add
			{
				CreateEventBridge();
				_NewDocumentEvent += value;
			}
			remove
			{
				_NewDocumentEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WindowActivateEventHandler _WindowActivateEvent;

		/// <summary>
		/// SupportByVersion Word 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840337.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public event Application_WindowActivateEventHandler WindowActivateEvent
		{
			add
			{
				CreateEventBridge();
				_WindowActivateEvent += value;
			}
			remove
			{
				_WindowActivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WindowDeactivateEventHandler _WindowDeactivateEvent;

		/// <summary>
		/// SupportByVersion Word 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198272.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public event Application_WindowDeactivateEventHandler WindowDeactivateEvent
		{
			add
			{
				CreateEventBridge();
				_WindowDeactivateEvent += value;
			}
			remove
			{
				_WindowDeactivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WindowSelectionChangeEventHandler _WindowSelectionChangeEvent;

		/// <summary>
		/// SupportByVersion Word 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192791.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public event Application_WindowSelectionChangeEventHandler WindowSelectionChangeEvent
		{
			add
			{
				CreateEventBridge();
				_WindowSelectionChangeEvent += value;
			}
			remove
			{
				_WindowSelectionChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WindowBeforeRightClickEventHandler _WindowBeforeRightClickEvent;

		/// <summary>
		/// SupportByVersion Word 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837868.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public event Application_WindowBeforeRightClickEventHandler WindowBeforeRightClickEvent
		{
			add
			{
				CreateEventBridge();
				_WindowBeforeRightClickEvent += value;
			}
			remove
			{
				_WindowBeforeRightClickEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WindowBeforeDoubleClickEventHandler _WindowBeforeDoubleClickEvent;

		/// <summary>
		/// SupportByVersion Word 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840048.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public event Application_WindowBeforeDoubleClickEventHandler WindowBeforeDoubleClickEvent
		{
			add
			{
				CreateEventBridge();
				_WindowBeforeDoubleClickEvent += value;
			}
			remove
			{
				_WindowBeforeDoubleClickEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 10,11,12,14,15,16
		/// </summary>
		private event Application_EPostagePropertyDialogEventHandler _EPostagePropertyDialogEvent;

		/// <summary>
		/// SupportByVersion Word 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197984.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public event Application_EPostagePropertyDialogEventHandler EPostagePropertyDialogEvent
		{
			add
			{
				CreateEventBridge();
				_EPostagePropertyDialogEvent += value;
			}
			remove
			{
				_EPostagePropertyDialogEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 10,11,12,14,15,16
		/// </summary>
		private event Application_EPostageInsertEventHandler _EPostageInsertEvent;

		/// <summary>
		/// SupportByVersion Word 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193389.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public event Application_EPostageInsertEventHandler EPostageInsertEvent
		{
			add
			{
				CreateEventBridge();
				_EPostageInsertEvent += value;
			}
			remove
			{
				_EPostageInsertEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 10,11,12,14,15,16
		/// </summary>
		private event Application_MailMergeAfterMergeEventHandler _MailMergeAfterMergeEvent;

		/// <summary>
		/// SupportByVersion Word 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198141.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public event Application_MailMergeAfterMergeEventHandler MailMergeAfterMergeEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeAfterMergeEvent += value;
			}
			remove
			{
				_MailMergeAfterMergeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 10,11,12,14,15,16
		/// </summary>
		private event Application_MailMergeAfterRecordMergeEventHandler _MailMergeAfterRecordMergeEvent;

		/// <summary>
		/// SupportByVersion Word 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198157.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public event Application_MailMergeAfterRecordMergeEventHandler MailMergeAfterRecordMergeEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeAfterRecordMergeEvent += value;
			}
			remove
			{
				_MailMergeAfterRecordMergeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 10,11,12,14,15,16
		/// </summary>
		private event Application_MailMergeBeforeMergeEventHandler _MailMergeBeforeMergeEvent;

		/// <summary>
		/// SupportByVersion Word 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834588.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public event Application_MailMergeBeforeMergeEventHandler MailMergeBeforeMergeEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeBeforeMergeEvent += value;
			}
			remove
			{
				_MailMergeBeforeMergeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 10,11,12,14,15,16
		/// </summary>
		private event Application_MailMergeBeforeRecordMergeEventHandler _MailMergeBeforeRecordMergeEvent;

		/// <summary>
		/// SupportByVersion Word 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838357.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public event Application_MailMergeBeforeRecordMergeEventHandler MailMergeBeforeRecordMergeEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeBeforeRecordMergeEvent += value;
			}
			remove
			{
				_MailMergeBeforeRecordMergeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 10,11,12,14,15,16
		/// </summary>
		private event Application_MailMergeDataSourceLoadEventHandler _MailMergeDataSourceLoadEvent;

		/// <summary>
		/// SupportByVersion Word 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196096.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public event Application_MailMergeDataSourceLoadEventHandler MailMergeDataSourceLoadEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeDataSourceLoadEvent += value;
			}
			remove
			{
				_MailMergeDataSourceLoadEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 10,11,12,14,15,16
		/// </summary>
		private event Application_MailMergeDataSourceValidateEventHandler _MailMergeDataSourceValidateEvent;

		/// <summary>
		/// SupportByVersion Word 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193130.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public event Application_MailMergeDataSourceValidateEventHandler MailMergeDataSourceValidateEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeDataSourceValidateEvent += value;
			}
			remove
			{
				_MailMergeDataSourceValidateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 10,11,12,14,15,16
		/// </summary>
		private event Application_MailMergeWizardSendToCustomEventHandler _MailMergeWizardSendToCustomEvent;

		/// <summary>
		/// SupportByVersion Word 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837009.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public event Application_MailMergeWizardSendToCustomEventHandler MailMergeWizardSendToCustomEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeWizardSendToCustomEvent += value;
			}
			remove
			{
				_MailMergeWizardSendToCustomEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 10,11,12,14,15,16
		/// </summary>
		private event Application_MailMergeWizardStateChangeEventHandler _MailMergeWizardStateChangeEvent;

		/// <summary>
		/// SupportByVersion Word 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838546.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public event Application_MailMergeWizardStateChangeEventHandler MailMergeWizardStateChangeEvent
		{
			add
			{
				CreateEventBridge();
				_MailMergeWizardStateChangeEvent += value;
			}
			remove
			{
				_MailMergeWizardStateChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 10,11,12,14,15,16
		/// </summary>
		private event Application_WindowSizeEventHandler _WindowSizeEvent;

		/// <summary>
		/// SupportByVersion Word 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834597.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public event Application_WindowSizeEventHandler WindowSizeEvent
		{
			add
			{
				CreateEventBridge();
				_WindowSizeEvent += value;
			}
			remove
			{
				_WindowSizeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 11,12,14,15,16
		/// </summary>
		private event Application_XMLSelectionChangeEventHandler _XMLSelectionChangeEvent;

		/// <summary>
		/// SupportByVersion Word 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835495.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public event Application_XMLSelectionChangeEventHandler XMLSelectionChangeEvent
		{
			add
			{
				CreateEventBridge();
				_XMLSelectionChangeEvent += value;
			}
			remove
			{
				_XMLSelectionChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 11,12,14,15,16
		/// </summary>
		private event Application_XMLValidationErrorEventHandler _XMLValidationErrorEvent;

		/// <summary>
		/// SupportByVersion Word 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837452.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public event Application_XMLValidationErrorEventHandler XMLValidationErrorEvent
		{
			add
			{
				CreateEventBridge();
				_XMLValidationErrorEvent += value;
			}
			remove
			{
				_XMLValidationErrorEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 11,12,14,15,16
		/// </summary>
		private event Application_DocumentSyncEventHandler _DocumentSyncEvent;

		/// <summary>
		/// SupportByVersion Word 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835138.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public event Application_DocumentSyncEventHandler DocumentSyncEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentSyncEvent += value;
			}
			remove
			{
				_DocumentSyncEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 11,12,14,15,16
		/// </summary>
		private event Application_EPostageInsertExEventHandler _EPostageInsertExEvent;

		/// <summary>
		/// SupportByVersion Word 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195087.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public event Application_EPostageInsertExEventHandler EPostageInsertExEvent
		{
			add
			{
				CreateEventBridge();
				_EPostageInsertExEvent += value;
			}
			remove
			{
				_EPostageInsertExEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 12,14,15,16
		/// </summary>
		private event Application_MailMergeDataSourceValidate2EventHandler _MailMergeDataSourceValidate2Event;

		/// <summary>
		/// SupportByVersion Word 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839145.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public event Application_MailMergeDataSourceValidate2EventHandler MailMergeDataSourceValidate2Event
		{
			add
			{
				CreateEventBridge();
				_MailMergeDataSourceValidate2Event += value;
			}
			remove
			{
				_MailMergeDataSourceValidate2Event -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 14,15,16
		/// </summary>
		private event Application_ProtectedViewWindowOpenEventHandler _ProtectedViewWindowOpenEvent;

		/// <summary>
		/// SupportByVersion Word 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194483.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public event Application_ProtectedViewWindowOpenEventHandler ProtectedViewWindowOpenEvent
		{
			add
			{
				CreateEventBridge();
				_ProtectedViewWindowOpenEvent += value;
			}
			remove
			{
				_ProtectedViewWindowOpenEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 14,15,16
		/// </summary>
		private event Application_ProtectedViewWindowBeforeEditEventHandler _ProtectedViewWindowBeforeEditEvent;

		/// <summary>
		/// SupportByVersion Word 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192123.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public event Application_ProtectedViewWindowBeforeEditEventHandler ProtectedViewWindowBeforeEditEvent
		{
			add
			{
				CreateEventBridge();
				_ProtectedViewWindowBeforeEditEvent += value;
			}
			remove
			{
				_ProtectedViewWindowBeforeEditEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 14,15,16
		/// </summary>
		private event Application_ProtectedViewWindowBeforeCloseEventHandler _ProtectedViewWindowBeforeCloseEvent;

		/// <summary>
		/// SupportByVersion Word 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194718.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public event Application_ProtectedViewWindowBeforeCloseEventHandler ProtectedViewWindowBeforeCloseEvent
		{
			add
			{
				CreateEventBridge();
				_ProtectedViewWindowBeforeCloseEvent += value;
			}
			remove
			{
				_ProtectedViewWindowBeforeCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 14,15,16
		/// </summary>
		private event Application_ProtectedViewWindowSizeEventHandler _ProtectedViewWindowSizeEvent;

		/// <summary>
		/// SupportByVersion Word 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836722.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public event Application_ProtectedViewWindowSizeEventHandler ProtectedViewWindowSizeEvent
		{
			add
			{
				CreateEventBridge();
				_ProtectedViewWindowSizeEvent += value;
			}
			remove
			{
				_ProtectedViewWindowSizeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 14,15,16
		/// </summary>
		private event Application_ProtectedViewWindowActivateEventHandler _ProtectedViewWindowActivateEvent;

		/// <summary>
		/// SupportByVersion Word 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836396.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public event Application_ProtectedViewWindowActivateEventHandler ProtectedViewWindowActivateEvent
		{
			add
			{
				CreateEventBridge();
				_ProtectedViewWindowActivateEvent += value;
			}
			remove
			{
				_ProtectedViewWindowActivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Word, 14,15,16
		/// </summary>
		private event Application_ProtectedViewWindowDeactivateEventHandler _ProtectedViewWindowDeactivateEvent;

		/// <summary>
		/// SupportByVersion Word 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837500.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public event Application_ProtectedViewWindowDeactivateEventHandler ProtectedViewWindowDeactivateEvent
		{
			add
			{
				CreateEventBridge();
				_ProtectedViewWindowDeactivateEvent += value;
			}
			remove
			{
				_ProtectedViewWindowDeactivateEvent -= value;
			}
		}

		#endregion
       
	    #region IEventBinding
        
		/// <summary>
        /// Creates active sink helper
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void CreateEventBridge()
        {
			if(false == Factory.Settings.EnableEvents)
				return;
	
			if (null != _connectPoint)
				return;
	
            if (null == _activeSinkId)
				_activeSinkId = SinkHelper.GetConnectionPoint2(this, ref _connectPoint, Events.ApplicationEvents2_SinkHelper.Id, Events.ApplicationEvents3_SinkHelper.Id, Events.ApplicationEvents4_SinkHelper.Id);


			if(Events.ApplicationEvents2_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_applicationEvents2_SinkHelper = new Events.ApplicationEvents2_SinkHelper(this, _connectPoint);
				return;
			}

			if(Events.ApplicationEvents3_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_applicationEvents3_SinkHelper = new Events.ApplicationEvents3_SinkHelper(this, _connectPoint);
				return;
			}

			if(Events.ApplicationEvents4_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_applicationEvents4_SinkHelper = new Events.ApplicationEvents4_SinkHelper(this, _connectPoint);
				return;
			} 
        }

        /// <summary>
        /// The instance use currently an event listener 
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool EventBridgeInitialized
        {
            get 
            {
                return (null != _connectPoint);
            }
        }
        /// <summary>
        /// Instance has one or more event recipients
        /// </summary>
        /// <returns>true if one or more event is active, otherwise false</returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool HasEventRecipients()       
        {
            return NetOffice.Events.CoClassEventReflector.HasEventRecipients(this, LateBindingApiWrapperType);            
        }

        /// <summary>
        /// Instance has one or more event recipients
        /// </summary>
        /// <param name="eventName">name of the event</param>
        /// <returns></returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool HasEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.HasEventRecipients(this, LateBindingApiWrapperType, eventName);
        }

        /// <summary>
        /// Target methods from its actual event recipients
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Delegate[] GetEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.GetEventRecipients(this, LateBindingApiWrapperType, eventName);
        }
       
        /// <summary>
        /// Returns the current count of event recipients
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public int GetCountOfEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.GetCountOfEventRecipients(this, LateBindingApiWrapperType, eventName);       
         }
        
        /// <summary>
        /// Raise an instance event
        /// </summary>
        /// <param name="eventName">name of the event without 'Event' at the end</param>
        /// <param name="paramsArray">custom arguments for the event</param>
        /// <returns>count of called event recipients</returns>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public int RaiseCustomEvent(string eventName, ref object[] paramsArray)
		{
            return NetOffice.Events.CoClassEventReflector.RaiseCustomEvent(this, LateBindingApiWrapperType, eventName, ref paramsArray);
		}
        /// <summary>
        /// Stop listening events for the instance
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void DisposeEventBridge()
        {
			if( null != _applicationEvents2_SinkHelper)
			{
				_applicationEvents2_SinkHelper.Dispose();
				_applicationEvents2_SinkHelper = null;
			}
			if( null != _applicationEvents3_SinkHelper)
			{
				_applicationEvents3_SinkHelper.Dispose();
				_applicationEvents3_SinkHelper = null;
			}
			if( null != _applicationEvents4_SinkHelper)
			{
				_applicationEvents4_SinkHelper.Dispose();
				_applicationEvents4_SinkHelper = null;
			}

			_connectPoint = null;
		}

        #endregion

        #region ICloneable<Application>

        /// <summary>
        /// Creates a new Application that is a copy of the current instance
        /// </summary>
        /// <returns>A new Application that is a copy of this instance</returns>
        /// <exception cref="CloneException">An unexpected error occured. See inner exception(s) for details.</exception>
        public new virtual Application Clone()
        {
            return base.Clone() as Application;
        }

        #endregion

        #pragma warning restore
    }
}