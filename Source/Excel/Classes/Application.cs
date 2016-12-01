using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
using NetOffice.Misc;

namespace NetOffice.ExcelApi
{

	#region Delegates

	#pragma warning disable
	public delegate void Application_NewWorkbookEventHandler(NetOffice.ExcelApi.Workbook Wb);
	public delegate void Application_SheetSelectionChangeEventHandler(COMObject Sh, NetOffice.ExcelApi.Range Target);
	public delegate void Application_SheetBeforeDoubleClickEventHandler(COMObject Sh, NetOffice.ExcelApi.Range Target, ref bool Cancel);
	public delegate void Application_SheetBeforeRightClickEventHandler(COMObject Sh, NetOffice.ExcelApi.Range Target, ref bool Cancel);
	public delegate void Application_SheetActivateEventHandler(COMObject Sh);
	public delegate void Application_SheetDeactivateEventHandler(COMObject Sh);
	public delegate void Application_SheetCalculateEventHandler(COMObject Sh);
	public delegate void Application_SheetChangeEventHandler(COMObject Sh, NetOffice.ExcelApi.Range Target);
	public delegate void Application_WorkbookOpenEventHandler(NetOffice.ExcelApi.Workbook Wb);
	public delegate void Application_WorkbookActivateEventHandler(NetOffice.ExcelApi.Workbook Wb);
	public delegate void Application_WorkbookDeactivateEventHandler(NetOffice.ExcelApi.Workbook Wb);
	public delegate void Application_WorkbookBeforeCloseEventHandler(NetOffice.ExcelApi.Workbook Wb, ref bool Cancel);
	public delegate void Application_WorkbookBeforeSaveEventHandler(NetOffice.ExcelApi.Workbook Wb, bool SaveAsUI, ref bool Cancel);
	public delegate void Application_WorkbookBeforePrintEventHandler(NetOffice.ExcelApi.Workbook Wb, ref bool Cancel);
	public delegate void Application_WorkbookNewSheetEventHandler(NetOffice.ExcelApi.Workbook Wb, COMObject Sh);
	public delegate void Application_WorkbookAddinInstallEventHandler(NetOffice.ExcelApi.Workbook Wb);
	public delegate void Application_WorkbookAddinUninstallEventHandler(NetOffice.ExcelApi.Workbook Wb);
	public delegate void Application_WindowResizeEventHandler(NetOffice.ExcelApi.Workbook Wb, NetOffice.ExcelApi.Window Wn);
	public delegate void Application_WindowActivateEventHandler(NetOffice.ExcelApi.Workbook Wb, NetOffice.ExcelApi.Window Wn);
	public delegate void Application_WindowDeactivateEventHandler(NetOffice.ExcelApi.Workbook Wb, NetOffice.ExcelApi.Window Wn);
	public delegate void Application_SheetFollowHyperlinkEventHandler(COMObject Sh, NetOffice.ExcelApi.Hyperlink Target);
	public delegate void Application_SheetPivotTableUpdateEventHandler(COMObject Sh, NetOffice.ExcelApi.PivotTable Target);
	public delegate void Application_WorkbookPivotTableCloseConnectionEventHandler(NetOffice.ExcelApi.Workbook Wb, NetOffice.ExcelApi.PivotTable Target);
	public delegate void Application_WorkbookPivotTableOpenConnectionEventHandler(NetOffice.ExcelApi.Workbook Wb, NetOffice.ExcelApi.PivotTable Target);
	public delegate void Application_WorkbookSyncEventHandler(NetOffice.ExcelApi.Workbook Wb, NetOffice.OfficeApi.Enums.MsoSyncEventType SyncEventType);
	public delegate void Application_WorkbookBeforeXmlImportEventHandler(NetOffice.ExcelApi.Workbook Wb, NetOffice.ExcelApi.XmlMap Map, string Url, bool IsRefresh, ref bool Cancel);
	public delegate void Application_WorkbookAfterXmlImportEventHandler(NetOffice.ExcelApi.Workbook Wb, NetOffice.ExcelApi.XmlMap Map, bool IsRefresh, NetOffice.ExcelApi.Enums.XlXmlImportResult Result);
	public delegate void Application_WorkbookBeforeXmlExportEventHandler(NetOffice.ExcelApi.Workbook Wb, NetOffice.ExcelApi.XmlMap Map, string Url, ref bool Cancel);
	public delegate void Application_WorkbookAfterXmlExportEventHandler(NetOffice.ExcelApi.Workbook Wb, NetOffice.ExcelApi.XmlMap Map, string Url, NetOffice.ExcelApi.Enums.XlXmlExportResult Result);
	public delegate void Application_WorkbookRowsetCompleteEventHandler(NetOffice.ExcelApi.Workbook Wb, string Description, string Sheet, bool Success);
	public delegate void Application_AfterCalculateEventHandler();
	public delegate void Application_SheetPivotTableAfterValueChangeEventHandler(COMObject Sh, NetOffice.ExcelApi.PivotTable TargetPivotTable, NetOffice.ExcelApi.Range TargetRange);
	public delegate void Application_SheetPivotTableBeforeAllocateChangesEventHandler(COMObject Sh, NetOffice.ExcelApi.PivotTable TargetPivotTable, Int32 ValueChangeStart, Int32 ValueChangeEnd, ref bool Cancel);
	public delegate void Application_SheetPivotTableBeforeCommitChangesEventHandler(COMObject Sh, NetOffice.ExcelApi.PivotTable TargetPivotTable, Int32 ValueChangeStart, Int32 ValueChangeEnd, ref bool Cancel);
	public delegate void Application_SheetPivotTableBeforeDiscardChangesEventHandler(COMObject Sh, NetOffice.ExcelApi.PivotTable TargetPivotTable, Int32 ValueChangeStart, Int32 ValueChangeEnd);
	public delegate void Application_ProtectedViewWindowOpenEventHandler(NetOffice.ExcelApi.ProtectedViewWindow Pvw);
	public delegate void Application_ProtectedViewWindowBeforeEditEventHandler(NetOffice.ExcelApi.ProtectedViewWindow Pvw, ref bool Cancel);
	public delegate void Application_ProtectedViewWindowBeforeCloseEventHandler(NetOffice.ExcelApi.ProtectedViewWindow Pvw, NetOffice.ExcelApi.Enums.XlProtectedViewCloseReason Reason, ref bool Cancel);
	public delegate void Application_ProtectedViewWindowResizeEventHandler(NetOffice.ExcelApi.ProtectedViewWindow Pvw);
	public delegate void Application_ProtectedViewWindowActivateEventHandler(NetOffice.ExcelApi.ProtectedViewWindow Pvw);
	public delegate void Application_ProtectedViewWindowDeactivateEventHandler(NetOffice.ExcelApi.ProtectedViewWindow Pvw);
	public delegate void Application_WorkbookAfterSaveEventHandler(NetOffice.ExcelApi.Workbook Wb, bool Success);
	public delegate void Application_WorkbookNewChartEventHandler(NetOffice.ExcelApi.Workbook Wb, NetOffice.ExcelApi.Chart Ch);
	public delegate void Application_SheetLensGalleryRenderCompleteEventHandler(COMObject Sh);
	public delegate void Application_SheetTableUpdateEventHandler(COMObject Sh, NetOffice.ExcelApi.TableObject Target);
	public delegate void Application_WorkbookModelChangeEventHandler(NetOffice.ExcelApi.Workbook Wb, NetOffice.ExcelApi.ModelChanges Changes);
	public delegate void Application_SheetBeforeDeleteEventHandler(COMObject Sh);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass Application 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194565.aspx
	///</summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class Application : _Application,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		AppEvents_SinkHelper _appEvents_SinkHelper;
	
		#endregion

		#region Type Information

        /// <summary>
        /// Instance Type
        /// </summary>
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
                    _type = typeof(Application);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

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
		
		///<summary>
        /// Creates a new instance of Application 
        ///</summary>		
		public Application():base("Excel.Application")
		{
			_callQuitInDispose = true;
			GlobalHelperModules.GlobalModule.Instance = this;
		}
		
		///<summary>
        /// Creates a new instance of Application
        ///</summary>
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
		public override void Dispose(bool disposeEventBinding)
		{
			if(this.Equals(GlobalHelperModules.GlobalModule.Instance))
				 GlobalHelperModules.GlobalModule.Instance = null;	
			base.Dispose(disposeEventBinding);
		}

		/// <summary>
		/// NetOffice method: dispose instance and all child instances
		/// </summary>
		public override void Dispose()
		{
			if(this.Equals(GlobalHelperModules.GlobalModule.Instance))
				 GlobalHelperModules.GlobalModule.Instance = null;
			base.Dispose();
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// Returns all running Excel.Application objects from the environment/system
        /// </summary>
        /// <returns>an Excel.Application array</returns>
		public static NetOffice.ExcelApi.Application[] GetActiveInstances()
		{		
			IDisposableEnumeration proxyList = NetOffice.ProxyService.GetActiveInstances("Excel","Application");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.ExcelApi.Application> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.ExcelApi.Application>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.ExcelApi.Application(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// Returns a running Excel.Application object from the environment/system.
        /// </summary>
        /// <returns>an Excel.Application object or null</returns>
		public static NetOffice.ExcelApi.Application GetActiveInstance()
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("Excel","Application", false);
			if(null != proxy)
				return new NetOffice.ExcelApi.Application(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// Returns a running Excel.Application object from the environment/system. 
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Excel.Application object or null</returns>
		public static NetOffice.ExcelApi.Application GetActiveInstance(bool throwOnError)
		{
			object proxy  = NetOffice.ProxyService.GetActiveInstance("Excel","Application", throwOnError);
			if(null != proxy)
				return new NetOffice.ExcelApi.Application(null, proxy);
			else
				return null;
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
       
	    #region IEventBinding Member
        
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, AppEvents_SinkHelper.Id);


			if(AppEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_appEvents_SinkHelper = new AppEvents_SinkHelper(this, _connectPoint);
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
        ///  The instance has currently one or more event recipients 
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool HasEventRecipients()       
        {
			if(null == _thisType)
				_thisType = this.GetType();
					
			foreach (NetRuntimeSystem.Reflection.EventInfo item in _thisType.GetEvents())
			{
				MulticastDelegate eventDelegate = (MulticastDelegate) _thisType.GetType().GetField(item.Name, 
																			NetRuntimeSystem.Reflection.BindingFlags.NonPublic |
																			NetRuntimeSystem.Reflection.BindingFlags.Instance).GetValue(this);
					
				if( (null != eventDelegate) && (eventDelegate.GetInvocationList().Length > 0) )
					return false;
			}
				
			return false;
        }
        
        /// <summary>
        /// Target methods from its actual event recipients
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Delegate[] GetEventRecipients(string eventName)
        {
			if(null == _thisType)
				_thisType = this.GetType();
             
            MulticastDelegate eventDelegate = (MulticastDelegate)_thisType.GetField(
                                                "_" + eventName + "Event",
                                                NetRuntimeSystem.Reflection.BindingFlags.Instance |
                                                NetRuntimeSystem.Reflection.BindingFlags.NonPublic).GetValue(this);

            if (null != eventDelegate)
            {
                Delegate[] delegates = eventDelegate.GetInvocationList();
                return delegates;
            }
            else
                return new Delegate[0];
        }
       
        /// <summary>
        /// Returns the current count of event recipients
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public int GetCountOfEventRecipients(string eventName)
        {
			if(null == _thisType)
				_thisType = this.GetType();
             
            MulticastDelegate eventDelegate = (MulticastDelegate)_thisType.GetField(
                                                "_" + eventName + "Event",
                                                NetRuntimeSystem.Reflection.BindingFlags.Instance |
                                                NetRuntimeSystem.Reflection.BindingFlags.NonPublic).GetValue(this);

            if (null != eventDelegate)
            {
                Delegate[] delegates = eventDelegate.GetInvocationList();
                return delegates.Length;
            }
            else
                return 0;
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
			if(null == _thisType)
				_thisType = this.GetType();
             
            MulticastDelegate eventDelegate = (MulticastDelegate)_thisType.GetField(
                                                "_" + eventName + "Event",
                                                NetRuntimeSystem.Reflection.BindingFlags.Instance |
                                                NetRuntimeSystem.Reflection.BindingFlags.NonPublic).GetValue(this);

            if (null != eventDelegate)
            {
                Delegate[] delegates = eventDelegate.GetInvocationList();
                foreach (var item in delegates)
                {
                    try
                    {
                        item.Method.Invoke(item.Target, paramsArray);
                    }
                    catch (NetRuntimeSystem.Exception exception)
                    {
                        Factory.Console.WriteException(exception);
                    }
                }
                return delegates.Length;
            }
            else
                return 0;
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

		#pragma warning restore
	}
}