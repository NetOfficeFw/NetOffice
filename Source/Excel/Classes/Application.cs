using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ExcelApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Application_NewWorkbookEventHandler(NetOffice.ExcelApi.Workbook wb);
	public delegate void Application_SheetSelectionChangeEventHandler(ICOMObject sh, NetOffice.ExcelApi.Range target);
	public delegate void Application_SheetBeforeDoubleClickEventHandler(ICOMObject sh, NetOffice.ExcelApi.Range target, ref bool cancel);
	public delegate void Application_SheetBeforeRightClickEventHandler(ICOMObject sh, NetOffice.ExcelApi.Range target, ref bool cancel);
	public delegate void Application_SheetActivateEventHandler(ICOMObject sh);
	public delegate void Application_SheetDeactivateEventHandler(ICOMObject sh);
	public delegate void Application_SheetCalculateEventHandler(ICOMObject sh);
	public delegate void Application_SheetChangeEventHandler(ICOMObject sh, NetOffice.ExcelApi.Range target);
	public delegate void Application_WorkbookOpenEventHandler(NetOffice.ExcelApi.Workbook wb);
	public delegate void Application_WorkbookActivateEventHandler(NetOffice.ExcelApi.Workbook wb);
	public delegate void Application_WorkbookDeactivateEventHandler(NetOffice.ExcelApi.Workbook wb);
	public delegate void Application_WorkbookBeforeCloseEventHandler(NetOffice.ExcelApi.Workbook wb, ref bool cancel);
	public delegate void Application_WorkbookBeforeSaveEventHandler(NetOffice.ExcelApi.Workbook wb, bool saveAsUI, ref bool cancel);
	public delegate void Application_WorkbookBeforePrintEventHandler(NetOffice.ExcelApi.Workbook wb, ref bool cancel);
	public delegate void Application_WorkbookNewSheetEventHandler(NetOffice.ExcelApi.Workbook wb, ICOMObject sh);
	public delegate void Application_WorkbookAddinInstallEventHandler(NetOffice.ExcelApi.Workbook wb);
	public delegate void Application_WorkbookAddinUninstallEventHandler(NetOffice.ExcelApi.Workbook wb);
	public delegate void Application_WindowResizeEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Window wn);
	public delegate void Application_WindowActivateEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Window wn);
	public delegate void Application_WindowDeactivateEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Window wn);
	public delegate void Application_SheetFollowHyperlinkEventHandler(ICOMObject sh, NetOffice.ExcelApi.Hyperlink target);
	public delegate void Application_SheetPivotTableUpdateEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable target);
	public delegate void Application_WorkbookPivotTableCloseConnectionEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.PivotTable target);
	public delegate void Application_WorkbookPivotTableOpenConnectionEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.PivotTable target);
	public delegate void Application_WorkbookSyncEventHandler(NetOffice.ExcelApi.Workbook Wb, NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType);
	public delegate void Application_WorkbookBeforeXmlImportEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, string url, bool isRefresh, ref bool cancel);
	public delegate void Application_WorkbookAfterXmlImportEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, bool isRefresh, NetOffice.ExcelApi.Enums.XlXmlImportResult result);
	public delegate void Application_WorkbookBeforeXmlExportEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, string url, ref bool cancel);
	public delegate void Application_WorkbookAfterXmlExportEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.XmlMap map, string url, NetOffice.ExcelApi.Enums.XlXmlExportResult result);
	public delegate void Application_WorkbookRowsetCompleteEventHandler(NetOffice.ExcelApi.Workbook wb, string description, string sheet, bool success);
	public delegate void Application_AfterCalculateEventHandler();
	public delegate void Application_SheetPivotTableAfterValueChangeEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, NetOffice.ExcelApi.Range targetRange);
	public delegate void Application_SheetPivotTableBeforeAllocateChangesEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, ref bool cancel);
	public delegate void Application_SheetPivotTableBeforeCommitChangesEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, ref bool cancel);
	public delegate void Application_SheetPivotTableBeforeDiscardChangesEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd);
	public delegate void Application_ProtectedViewWindowOpenEventHandler(NetOffice.ExcelApi.ProtectedViewWindow pvw);
	public delegate void Application_ProtectedViewWindowBeforeEditEventHandler(NetOffice.ExcelApi.ProtectedViewWindow pvw, ref bool cancel);
	public delegate void Application_ProtectedViewWindowBeforeCloseEventHandler(NetOffice.ExcelApi.ProtectedViewWindow pvw, NetOffice.ExcelApi.Enums.XlProtectedViewCloseReason reason, ref bool cancel);
	public delegate void Application_ProtectedViewWindowResizeEventHandler(NetOffice.ExcelApi.ProtectedViewWindow pvw);
	public delegate void Application_ProtectedViewWindowActivateEventHandler(NetOffice.ExcelApi.ProtectedViewWindow pvw);
	public delegate void Application_ProtectedViewWindowDeactivateEventHandler(NetOffice.ExcelApi.ProtectedViewWindow pvw);
	public delegate void Application_WorkbookAfterSaveEventHandler(NetOffice.ExcelApi.Workbook wb, bool success);
	public delegate void Application_WorkbookNewChartEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.Chart ch);
	public delegate void Application_SheetLensGalleryRenderCompleteEventHandler(ICOMObject sh);
	public delegate void Application_SheetTableUpdateEventHandler(ICOMObject sh, NetOffice.ExcelApi.TableObject target);
	public delegate void Application_WorkbookModelChangeEventHandler(NetOffice.ExcelApi.Workbook wb, NetOffice.ExcelApi.ModelChanges changes);
	public delegate void Application_SheetBeforeDeleteEventHandler(ICOMObject sh);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Application 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194565.aspx </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass), ComProgId("Excel.Application"), ModuleProvider(typeof(GlobalHelperModules.GlobalModule))]
	[EventSink(typeof(Events.AppEvents_SinkHelper))]
    [ComEventInterface(typeof(Events.AppEvents))]
    public class Application : _Application, ICloneable<Application>, IEventBinding
	{
        #pragma warning disable

        #region Fields

        private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events.AppEvents_SinkHelper _appEvents_SinkHelper;
	
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
        ///<param name="progId">registered ProgID</param>
		public Application(string progId):base(progId)
		{
			_callQuitInDispose = true;
			GlobalHelperModules.GlobalModule.Instance = this;
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
                object proxy = Running.ProxyService.GetActiveInstance("Excel", "Application", false);
                if (null != proxy)
                {
                    CreateFromProxy(proxy, true);
                    FromProxyService = true;
                }
                else
                {
                    CreateFromProgId("Excel.Application", true);
                }
            }
            else
            {
                CreateFromProgId("Excel.Application", true);
            }

            OnCreate();
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
        /// Returns all running Excel.Application instances from the environment/system
        /// </summary>
        /// <returns>Excel.Application sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances()
        {
            return Running.ProxyService.GetActiveInstances<Application>("Excel", "Application");
        }

        /// <summary>
        /// Returns a running Excel.Application instance from the environment/system
        /// </summary>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>Excel.Application instance or null</returns>
        public static Application GetActiveInstance(bool throwExceptionIfNotFound = false)
        {
            return Running.ProxyService.GetActiveInstance<Application>("Excel", "Application", throwExceptionIfNotFound);
        }

        #endregion

        #region Events

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        private event Application_NewWorkbookEventHandler _NewWorkbookEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837373.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Application_NewWorkbookEventHandler NewWorkbookEvent
		{
			add
			{
				CreateEventBridge();
				_NewWorkbookEvent += value;
			}
			remove
			{
				_NewWorkbookEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_SheetSelectionChangeEventHandler _SheetSelectionChangeEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839035.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Application_SheetSelectionChangeEventHandler SheetSelectionChangeEvent
		{
			add
			{
				CreateEventBridge();
				_SheetSelectionChangeEvent += value;
			}
			remove
			{
				_SheetSelectionChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_SheetBeforeDoubleClickEventHandler _SheetBeforeDoubleClickEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836225.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Application_SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClickEvent
		{
			add
			{
				CreateEventBridge();
				_SheetBeforeDoubleClickEvent += value;
			}
			remove
			{
				_SheetBeforeDoubleClickEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_SheetBeforeRightClickEventHandler _SheetBeforeRightClickEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840532.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Application_SheetBeforeRightClickEventHandler SheetBeforeRightClickEvent
		{
			add
			{
				CreateEventBridge();
				_SheetBeforeRightClickEvent += value;
			}
			remove
			{
				_SheetBeforeRightClickEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_SheetActivateEventHandler _SheetActivateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193288.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Application_SheetActivateEventHandler SheetActivateEvent
		{
			add
			{
				CreateEventBridge();
				_SheetActivateEvent += value;
			}
			remove
			{
				_SheetActivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_SheetDeactivateEventHandler _SheetDeactivateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823120.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Application_SheetDeactivateEventHandler SheetDeactivateEvent
		{
			add
			{
				CreateEventBridge();
				_SheetDeactivateEvent += value;
			}
			remove
			{
				_SheetDeactivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_SheetCalculateEventHandler _SheetCalculateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835607.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Application_SheetCalculateEventHandler SheetCalculateEvent
		{
			add
			{
				CreateEventBridge();
				_SheetCalculateEvent += value;
			}
			remove
			{
				_SheetCalculateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_SheetChangeEventHandler _SheetChangeEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193591.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Application_SheetChangeEventHandler SheetChangeEvent
		{
			add
			{
				CreateEventBridge();
				_SheetChangeEvent += value;
			}
			remove
			{
				_SheetChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WorkbookOpenEventHandler _WorkbookOpenEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196583.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Application_WorkbookOpenEventHandler WorkbookOpenEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookOpenEvent += value;
			}
			remove
			{
				_WorkbookOpenEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WorkbookActivateEventHandler _WorkbookActivateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837347.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Application_WorkbookActivateEventHandler WorkbookActivateEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookActivateEvent += value;
			}
			remove
			{
				_WorkbookActivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WorkbookDeactivateEventHandler _WorkbookDeactivateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193560.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Application_WorkbookDeactivateEventHandler WorkbookDeactivateEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookDeactivateEvent += value;
			}
			remove
			{
				_WorkbookDeactivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WorkbookBeforeCloseEventHandler _WorkbookBeforeCloseEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836770.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Application_WorkbookBeforeCloseEventHandler WorkbookBeforeCloseEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookBeforeCloseEvent += value;
			}
			remove
			{
				_WorkbookBeforeCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WorkbookBeforeSaveEventHandler _WorkbookBeforeSaveEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840422.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Application_WorkbookBeforeSaveEventHandler WorkbookBeforeSaveEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookBeforeSaveEvent += value;
			}
			remove
			{
				_WorkbookBeforeSaveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WorkbookBeforePrintEventHandler _WorkbookBeforePrintEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195507.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Application_WorkbookBeforePrintEventHandler WorkbookBeforePrintEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookBeforePrintEvent += value;
			}
			remove
			{
				_WorkbookBeforePrintEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WorkbookNewSheetEventHandler _WorkbookNewSheetEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198367.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Application_WorkbookNewSheetEventHandler WorkbookNewSheetEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookNewSheetEvent += value;
			}
			remove
			{
				_WorkbookNewSheetEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WorkbookAddinInstallEventHandler _WorkbookAddinInstallEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836206.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Application_WorkbookAddinInstallEventHandler WorkbookAddinInstallEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookAddinInstallEvent += value;
			}
			remove
			{
				_WorkbookAddinInstallEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WorkbookAddinUninstallEventHandler _WorkbookAddinUninstallEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835570.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Application_WorkbookAddinUninstallEventHandler WorkbookAddinUninstallEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookAddinUninstallEvent += value;
			}
			remove
			{
				_WorkbookAddinUninstallEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WindowResizeEventHandler _WindowResizeEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836166.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Application_WindowResizeEventHandler WindowResizeEvent
		{
			add
			{
				CreateEventBridge();
				_WindowResizeEvent += value;
			}
			remove
			{
				_WindowResizeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WindowActivateEventHandler _WindowActivateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821328.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_WindowDeactivateEventHandler _WindowDeactivateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822473.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Application_SheetFollowHyperlinkEventHandler _SheetFollowHyperlinkEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821956.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Application_SheetFollowHyperlinkEventHandler SheetFollowHyperlinkEvent
		{
			add
			{
				CreateEventBridge();
				_SheetFollowHyperlinkEvent += value;
			}
			remove
			{
				_SheetFollowHyperlinkEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 10,11,12,14,15,16
		/// </summary>
		private event Application_SheetPivotTableUpdateEventHandler _SheetPivotTableUpdateEvent;

		/// <summary>
		/// SupportByVersion Excel 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840950.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public event Application_SheetPivotTableUpdateEventHandler SheetPivotTableUpdateEvent
		{
			add
			{
				CreateEventBridge();
				_SheetPivotTableUpdateEvent += value;
			}
			remove
			{
				_SheetPivotTableUpdateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 10,11,12,14,15,16
		/// </summary>
		private event Application_WorkbookPivotTableCloseConnectionEventHandler _WorkbookPivotTableCloseConnectionEvent;

		/// <summary>
		/// SupportByVersion Excel 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198029.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public event Application_WorkbookPivotTableCloseConnectionEventHandler WorkbookPivotTableCloseConnectionEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookPivotTableCloseConnectionEvent += value;
			}
			remove
			{
				_WorkbookPivotTableCloseConnectionEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 10,11,12,14,15,16
		/// </summary>
		private event Application_WorkbookPivotTableOpenConnectionEventHandler _WorkbookPivotTableOpenConnectionEvent;

		/// <summary>
		/// SupportByVersion Excel 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821547.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public event Application_WorkbookPivotTableOpenConnectionEventHandler WorkbookPivotTableOpenConnectionEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookPivotTableOpenConnectionEvent += value;
			}
			remove
			{
				_WorkbookPivotTableOpenConnectionEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 11,12,14,15,16
		/// </summary>
		private event Application_WorkbookSyncEventHandler _WorkbookSyncEvent;

		/// <summary>
		/// SupportByVersion Excel 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839042.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public event Application_WorkbookSyncEventHandler WorkbookSyncEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookSyncEvent += value;
			}
			remove
			{
				_WorkbookSyncEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 11,12,14,15,16
		/// </summary>
		private event Application_WorkbookBeforeXmlImportEventHandler _WorkbookBeforeXmlImportEvent;

		/// <summary>
		/// SupportByVersion Excel 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196324.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public event Application_WorkbookBeforeXmlImportEventHandler WorkbookBeforeXmlImportEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookBeforeXmlImportEvent += value;
			}
			remove
			{
				_WorkbookBeforeXmlImportEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 11,12,14,15,16
		/// </summary>
		private event Application_WorkbookAfterXmlImportEventHandler _WorkbookAfterXmlImportEvent;

		/// <summary>
		/// SupportByVersion Excel 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837416.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public event Application_WorkbookAfterXmlImportEventHandler WorkbookAfterXmlImportEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookAfterXmlImportEvent += value;
			}
			remove
			{
				_WorkbookAfterXmlImportEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 11,12,14,15,16
		/// </summary>
		private event Application_WorkbookBeforeXmlExportEventHandler _WorkbookBeforeXmlExportEvent;

		/// <summary>
		/// SupportByVersion Excel 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195824.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public event Application_WorkbookBeforeXmlExportEventHandler WorkbookBeforeXmlExportEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookBeforeXmlExportEvent += value;
			}
			remove
			{
				_WorkbookBeforeXmlExportEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 11,12,14,15,16
		/// </summary>
		private event Application_WorkbookAfterXmlExportEventHandler _WorkbookAfterXmlExportEvent;

		/// <summary>
		/// SupportByVersion Excel 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836803.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public event Application_WorkbookAfterXmlExportEventHandler WorkbookAfterXmlExportEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookAfterXmlExportEvent += value;
			}
			remove
			{
				_WorkbookAfterXmlExportEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 12,14,15,16
		/// </summary>
		private event Application_WorkbookRowsetCompleteEventHandler _WorkbookRowsetCompleteEvent;

		/// <summary>
		/// SupportByVersion Excel 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839165.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public event Application_WorkbookRowsetCompleteEventHandler WorkbookRowsetCompleteEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookRowsetCompleteEvent += value;
			}
			remove
			{
				_WorkbookRowsetCompleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 12,14,15,16
		/// </summary>
		private event Application_AfterCalculateEventHandler _AfterCalculateEvent;

		/// <summary>
		/// SupportByVersion Excel 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840621.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public event Application_AfterCalculateEventHandler AfterCalculateEvent
		{
			add
			{
				CreateEventBridge();
				_AfterCalculateEvent += value;
			}
			remove
			{
				_AfterCalculateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Application_SheetPivotTableAfterValueChangeEventHandler _SheetPivotTableAfterValueChangeEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193316.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public event Application_SheetPivotTableAfterValueChangeEventHandler SheetPivotTableAfterValueChangeEvent
		{
			add
			{
				CreateEventBridge();
				_SheetPivotTableAfterValueChangeEvent += value;
			}
			remove
			{
				_SheetPivotTableAfterValueChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Application_SheetPivotTableBeforeAllocateChangesEventHandler _SheetPivotTableBeforeAllocateChangesEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838226.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public event Application_SheetPivotTableBeforeAllocateChangesEventHandler SheetPivotTableBeforeAllocateChangesEvent
		{
			add
			{
				CreateEventBridge();
				_SheetPivotTableBeforeAllocateChangesEvent += value;
			}
			remove
			{
				_SheetPivotTableBeforeAllocateChangesEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Application_SheetPivotTableBeforeCommitChangesEventHandler _SheetPivotTableBeforeCommitChangesEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838379.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public event Application_SheetPivotTableBeforeCommitChangesEventHandler SheetPivotTableBeforeCommitChangesEvent
		{
			add
			{
				CreateEventBridge();
				_SheetPivotTableBeforeCommitChangesEvent += value;
			}
			remove
			{
				_SheetPivotTableBeforeCommitChangesEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Application_SheetPivotTableBeforeDiscardChangesEventHandler _SheetPivotTableBeforeDiscardChangesEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835217.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public event Application_SheetPivotTableBeforeDiscardChangesEventHandler SheetPivotTableBeforeDiscardChangesEvent
		{
			add
			{
				CreateEventBridge();
				_SheetPivotTableBeforeDiscardChangesEvent += value;
			}
			remove
			{
				_SheetPivotTableBeforeDiscardChangesEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Application_ProtectedViewWindowOpenEventHandler _ProtectedViewWindowOpenEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194431.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
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
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Application_ProtectedViewWindowBeforeEditEventHandler _ProtectedViewWindowBeforeEditEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838239.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
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
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Application_ProtectedViewWindowBeforeCloseEventHandler _ProtectedViewWindowBeforeCloseEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821579.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
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
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Application_ProtectedViewWindowResizeEventHandler _ProtectedViewWindowResizeEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836848.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public event Application_ProtectedViewWindowResizeEventHandler ProtectedViewWindowResizeEvent
		{
			add
			{
				CreateEventBridge();
				_ProtectedViewWindowResizeEvent += value;
			}
			remove
			{
				_ProtectedViewWindowResizeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Application_ProtectedViewWindowActivateEventHandler _ProtectedViewWindowActivateEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195451.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
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
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Application_ProtectedViewWindowDeactivateEventHandler _ProtectedViewWindowDeactivateEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196820.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
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

		/// <summary>
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Application_WorkbookAfterSaveEventHandler _WorkbookAfterSaveEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198184.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public event Application_WorkbookAfterSaveEventHandler WorkbookAfterSaveEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookAfterSaveEvent += value;
			}
			remove
			{
				_WorkbookAfterSaveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Application_WorkbookNewChartEventHandler _WorkbookNewChartEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834985.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public event Application_WorkbookNewChartEventHandler WorkbookNewChartEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookNewChartEvent += value;
			}
			remove
			{
				_WorkbookNewChartEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 15, 16
		/// </summary>
		private event Application_SheetLensGalleryRenderCompleteEventHandler _SheetLensGalleryRenderCompleteEvent;

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227506.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public event Application_SheetLensGalleryRenderCompleteEventHandler SheetLensGalleryRenderCompleteEvent
		{
			add
			{
				CreateEventBridge();
				_SheetLensGalleryRenderCompleteEvent += value;
			}
			remove
			{
				_SheetLensGalleryRenderCompleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 15, 16
		/// </summary>
		private event Application_SheetTableUpdateEventHandler _SheetTableUpdateEvent;

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229805.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public event Application_SheetTableUpdateEventHandler SheetTableUpdateEvent
		{
			add
			{
				CreateEventBridge();
				_SheetTableUpdateEvent += value;
			}
			remove
			{
				_SheetTableUpdateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 15, 16
		/// </summary>
		private event Application_WorkbookModelChangeEventHandler _WorkbookModelChangeEvent;

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229611.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public event Application_WorkbookModelChangeEventHandler WorkbookModelChangeEvent
		{
			add
			{
				CreateEventBridge();
				_WorkbookModelChangeEvent += value;
			}
			remove
			{
				_WorkbookModelChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 15, 16
		/// </summary>
		private event Application_SheetBeforeDeleteEventHandler _SheetBeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/dn448391.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public event Application_SheetBeforeDeleteEventHandler SheetBeforeDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_SheetBeforeDeleteEvent += value;
			}
			remove
			{
				_SheetBeforeDeleteEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, Events.AppEvents_SinkHelper.Id);

			if(Events.AppEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_appEvents_SinkHelper = new Events.AppEvents_SinkHelper(this, _connectPoint);
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
			if( null != _appEvents_SinkHelper)
			{
				_appEvents_SinkHelper.Dispose();
				_appEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}

        #endregion

        #region ICloneable<Application>

        /// <summary>
        /// Creates a new Application that is a copy of the current instance
        /// </summary>
        /// <returns>A new Application that is a copy of this instance</returns>
        /// <exception cref="NetOffice.Exceptions.CloneException">An unexpected error occured. See inner exception(s) for details.</exception>
        public new virtual Application Clone()
        {
            return base.Clone() as Application;
        }

        #endregion

        #pragma warning restore
    }
}