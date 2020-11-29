﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ExcelApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Workbook_OpenEventHandler();
	public delegate void Workbook_ActivateEventHandler();
	public delegate void Workbook_DeactivateEventHandler();
	public delegate void Workbook_BeforeCloseEventHandler(ref bool cancel);
	public delegate void Workbook_BeforeSaveEventHandler(bool saveAsUI, ref bool cancel);
	public delegate void Workbook_BeforePrintEventHandler(ref bool cancel);
	public delegate void Workbook_NewSheetEventHandler(ICOMObject sh);
	public delegate void Workbook_AddinInstallEventHandler();
	public delegate void Workbook_AddinUninstallEventHandler();
	public delegate void Workbook_WindowResizeEventHandler(NetOffice.ExcelApi.Window wn);
	public delegate void Workbook_WindowActivateEventHandler(NetOffice.ExcelApi.Window wn);
	public delegate void Workbook_WindowDeactivateEventHandler(NetOffice.ExcelApi.Window wn);
	public delegate void Workbook_SheetSelectionChangeEventHandler(ICOMObject sh, NetOffice.ExcelApi.Range target);
	public delegate void Workbook_SheetBeforeDoubleClickEventHandler(ICOMObject sh, NetOffice.ExcelApi.Range target, ref bool cancel);
	public delegate void Workbook_SheetBeforeRightClickEventHandler(ICOMObject sh, NetOffice.ExcelApi.Range target, ref bool cancel);
	public delegate void Workbook_SheetActivateEventHandler(ICOMObject sh);
	public delegate void Workbook_SheetDeactivateEventHandler(ICOMObject sh);
	public delegate void Workbook_SheetCalculateEventHandler(ICOMObject sh);
	public delegate void Workbook_SheetChangeEventHandler(ICOMObject sh, NetOffice.ExcelApi.Range target);
	public delegate void Workbook_SheetFollowHyperlinkEventHandler(ICOMObject sh, NetOffice.ExcelApi.Hyperlink target);
	public delegate void Workbook_SheetPivotTableUpdateEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable target);
	public delegate void Workbook_PivotTableCloseConnectionEventHandler(NetOffice.ExcelApi.PivotTable target);
	public delegate void Workbook_PivotTableOpenConnectionEventHandler(NetOffice.ExcelApi.PivotTable target);
	public delegate void Workbook_SyncEventHandler(NetOffice.OfficeApi.Enums.MsoSyncEventType syncEventType);
	public delegate void Workbook_BeforeXmlImportEventHandler(NetOffice.ExcelApi.XmlMap map, string url, bool IsRefresh, ref bool cancel);
	public delegate void Workbook_AfterXmlImportEventHandler(NetOffice.ExcelApi.XmlMap map, bool isRefresh, NetOffice.ExcelApi.Enums.XlXmlImportResult result);
	public delegate void Workbook_BeforeXmlExportEventHandler(NetOffice.ExcelApi.XmlMap map, string url, ref bool cancel);
	public delegate void Workbook_AfterXmlExportEventHandler(NetOffice.ExcelApi.XmlMap map, string url, NetOffice.ExcelApi.Enums.XlXmlExportResult result);
	public delegate void Workbook_RowsetCompleteEventHandler(string description, string sheet, bool success);
	public delegate void Workbook_SheetPivotTableAfterValueChangeEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, NetOffice.ExcelApi.Range targetRange);
	public delegate void Workbook_SheetPivotTableBeforeAllocateChangesEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, ref bool cancel);
	public delegate void Workbook_SheetPivotTableBeforeCommitChangesEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, ref bool cancel);
	public delegate void Workbook_SheetPivotTableBeforeDiscardChangesEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd);
	public delegate void Workbook_SheetPivotTableChangeSyncEventHandler(ICOMObject sh, NetOffice.ExcelApi.PivotTable target);
	public delegate void Workbook_AfterSaveEventHandler(bool success);
	public delegate void Workbook_NewChartEventHandler(NetOffice.ExcelApi.Chart ch);
	public delegate void Workbook_SheetLensGalleryRenderCompleteEventHandler(ICOMObject sh);
	public delegate void Workbook_SheetTableUpdateEventHandler(ICOMObject sh, NetOffice.ExcelApi.TableObject target);
	public delegate void Workbook_ModelChangeEventHandler(NetOffice.ExcelApi.ModelChanges changes);
	public delegate void Workbook_SheetBeforeDeleteEventHandler(ICOMObject sh);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Workbook 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[EventSink(typeof(Events.WorkbookEvents_SinkHelper))]
    [ComEventInterface(typeof(Events.WorkbookEvents))]
    public class Workbook : _Workbook, IEventBinding
	{
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events.WorkbookEvents_SinkHelper _workbookEvents_SinkHelper;
	
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
                    _type = typeof(Workbook);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Workbook(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Workbook(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Workbook(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Workbook(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Workbook(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of Workbook 
        /// </summary>		
		public Workbook():base("Excel.Workbook")
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of Workbook
        /// </summary>
        ///<param name="progId">registered ProgID</param>
		public Workbook(string progId):base(progId)
		{
			
		}

        #endregion

        #region Static CoClass Methods

        /// <summary>
        /// Returns all running Excel.Workbook instances from the environment/system
        /// </summary>
        /// <returns>Excel.Workbook sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances()
        {
            return Running.ProxyService.GetActiveInstances<Application>("Excel", "Workbook");
        }

        /// <summary>
        /// Returns a running Excel.Workbook instance from the environment/system
        /// </summary>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>Excel.Workbook instance or null</returns>
        public static Application GetActiveInstance(bool throwExceptionIfNotFound = false)
        {
            return Running.ProxyService.GetActiveInstance<Application>("Excel", "Workbook", throwExceptionIfNotFound);
        }

        #endregion

        #region Events

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        private event Workbook_OpenEventHandler _OpenEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.Open"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Workbook_OpenEventHandler OpenEvent
		{
			add
			{
				CreateEventBridge();
				_OpenEvent += value;
			}
			remove
			{
				_OpenEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Workbook_ActivateEventHandler _ActivateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.Activate(even)"/> </remarks>
		[SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public event Workbook_ActivateEventHandler ActivateEvent
        {
            add
            {
                CreateEventBridge();
                _ActivateEvent += value;
            }
            remove
            {
                _ActivateEvent -= value;
            }
        }

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Workbook_DeactivateEventHandler _DeactivateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.Deactivate"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Workbook_DeactivateEventHandler DeactivateEvent
		{
			add
			{
				CreateEventBridge();
				_DeactivateEvent += value;
			}
			remove
			{
				_DeactivateEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Workbook_BeforeCloseEventHandler _BeforeCloseEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.BeforeClose"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Workbook_BeforeCloseEventHandler BeforeCloseEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeCloseEvent += value;
			}
			remove
			{
				_BeforeCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Workbook_BeforeSaveEventHandler _BeforeSaveEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.BeforeSave"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Workbook_BeforeSaveEventHandler BeforeSaveEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeSaveEvent += value;
			}
			remove
			{
				_BeforeSaveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Workbook_BeforePrintEventHandler _BeforePrintEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.BeforePrint"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Workbook_BeforePrintEventHandler BeforePrintEvent
		{
			add
			{
				CreateEventBridge();
				_BeforePrintEvent += value;
			}
			remove
			{
				_BeforePrintEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Workbook_NewSheetEventHandler _NewSheetEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.NewSheet"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Workbook_NewSheetEventHandler NewSheetEvent
		{
			add
			{
				CreateEventBridge();
				_NewSheetEvent += value;
			}
			remove
			{
				_NewSheetEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Workbook_AddinInstallEventHandler _AddinInstallEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.AddinInstall"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Workbook_AddinInstallEventHandler AddinInstallEvent
		{
			add
			{
				CreateEventBridge();
				_AddinInstallEvent += value;
			}
			remove
			{
				_AddinInstallEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Workbook_AddinUninstallEventHandler _AddinUninstallEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.AddinUninstall"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Workbook_AddinUninstallEventHandler AddinUninstallEvent
		{
			add
			{
				CreateEventBridge();
				_AddinUninstallEvent += value;
			}
			remove
			{
				_AddinUninstallEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		private event Workbook_WindowResizeEventHandler _WindowResizeEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.WindowResize"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Workbook_WindowResizeEventHandler WindowResizeEvent
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
		private event Workbook_WindowActivateEventHandler _WindowActivateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.WindowActivate"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Workbook_WindowActivateEventHandler WindowActivateEvent
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
		private event Workbook_WindowDeactivateEventHandler _WindowDeactivateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.WindowDeactivate"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Workbook_WindowDeactivateEventHandler WindowDeactivateEvent
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
		private event Workbook_SheetSelectionChangeEventHandler _SheetSelectionChangeEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.SheetSelectionChange"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Workbook_SheetSelectionChangeEventHandler SheetSelectionChangeEvent
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
		private event Workbook_SheetBeforeDoubleClickEventHandler _SheetBeforeDoubleClickEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.SheetBeforeDoubleClick"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Workbook_SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClickEvent
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
		private event Workbook_SheetBeforeRightClickEventHandler _SheetBeforeRightClickEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.SheetBeforeRightClick"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Workbook_SheetBeforeRightClickEventHandler SheetBeforeRightClickEvent
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
		private event Workbook_SheetActivateEventHandler _SheetActivateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.SheetActivate"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Workbook_SheetActivateEventHandler SheetActivateEvent
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
		private event Workbook_SheetDeactivateEventHandler _SheetDeactivateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.SheetDeactivate"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Workbook_SheetDeactivateEventHandler SheetDeactivateEvent
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
		private event Workbook_SheetCalculateEventHandler _SheetCalculateEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.SheetCalculate"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Workbook_SheetCalculateEventHandler SheetCalculateEvent
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
		private event Workbook_SheetChangeEventHandler _SheetChangeEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.SheetChange"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Workbook_SheetChangeEventHandler SheetChangeEvent
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
		private event Workbook_SheetFollowHyperlinkEventHandler _SheetFollowHyperlinkEvent;

		/// <summary>
		/// SupportByVersion Excel 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.SheetFollowHyperlink"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public event Workbook_SheetFollowHyperlinkEventHandler SheetFollowHyperlinkEvent
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
		private event Workbook_SheetPivotTableUpdateEventHandler _SheetPivotTableUpdateEvent;

		/// <summary>
		/// SupportByVersion Excel 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.SheetPivotTableUpdate"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public event Workbook_SheetPivotTableUpdateEventHandler SheetPivotTableUpdateEvent
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
		private event Workbook_PivotTableCloseConnectionEventHandler _PivotTableCloseConnectionEvent;

		/// <summary>
		/// SupportByVersion Excel 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.workbook.pivottablecloseconnection"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public event Workbook_PivotTableCloseConnectionEventHandler PivotTableCloseConnectionEvent
		{
			add
			{
				CreateEventBridge();
				_PivotTableCloseConnectionEvent += value;
			}
			remove
			{
				_PivotTableCloseConnectionEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 10,11,12,14,15,16
		/// </summary>
		private event Workbook_PivotTableOpenConnectionEventHandler _PivotTableOpenConnectionEvent;

		/// <summary>
		/// SupportByVersion Excel 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.workbook.pivottableopenconnection"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public event Workbook_PivotTableOpenConnectionEventHandler PivotTableOpenConnectionEvent
		{
			add
			{
				CreateEventBridge();
				_PivotTableOpenConnectionEvent += value;
			}
			remove
			{
				_PivotTableOpenConnectionEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 11,12,14,15,16
		/// </summary>
		private event Workbook_SyncEventHandler _SyncEvent;

		/// <summary>
		/// SupportByVersion Excel 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.workbook.sync(event)"/> </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public event Workbook_SyncEventHandler SyncEvent
		{
			add
			{
				CreateEventBridge();
				_SyncEvent += value;
			}
			remove
			{
				_SyncEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 11,12,14,15,16
		/// </summary>
		private event Workbook_BeforeXmlImportEventHandler _BeforeXmlImportEvent;

		/// <summary>
		/// SupportByVersion Excel 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.BeforeXmlImport"/> </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public event Workbook_BeforeXmlImportEventHandler BeforeXmlImportEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeXmlImportEvent += value;
			}
			remove
			{
				_BeforeXmlImportEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 11,12,14,15,16
		/// </summary>
		private event Workbook_AfterXmlImportEventHandler _AfterXmlImportEvent;

		/// <summary>
		/// SupportByVersion Excel 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.AfterXmlImport"/> </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public event Workbook_AfterXmlImportEventHandler AfterXmlImportEvent
		{
			add
			{
				CreateEventBridge();
				_AfterXmlImportEvent += value;
			}
			remove
			{
				_AfterXmlImportEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 11,12,14,15,16
		/// </summary>
		private event Workbook_BeforeXmlExportEventHandler _BeforeXmlExportEvent;

		/// <summary>
		/// SupportByVersion Excel 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.BeforeXmlExport"/> </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public event Workbook_BeforeXmlExportEventHandler BeforeXmlExportEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeXmlExportEvent += value;
			}
			remove
			{
				_BeforeXmlExportEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 11,12,14,15,16
		/// </summary>
		private event Workbook_AfterXmlExportEventHandler _AfterXmlExportEvent;

		/// <summary>
		/// SupportByVersion Excel 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.AfterXmlExport"/> </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public event Workbook_AfterXmlExportEventHandler AfterXmlExportEvent
		{
			add
			{
				CreateEventBridge();
				_AfterXmlExportEvent += value;
			}
			remove
			{
				_AfterXmlExportEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 12,14,15,16
		/// </summary>
		private event Workbook_RowsetCompleteEventHandler _RowsetCompleteEvent;

		/// <summary>
		/// SupportByVersion Excel 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.RowsetComplete"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public event Workbook_RowsetCompleteEventHandler RowsetCompleteEvent
		{
			add
			{
				CreateEventBridge();
				_RowsetCompleteEvent += value;
			}
			remove
			{
				_RowsetCompleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Workbook_SheetPivotTableAfterValueChangeEventHandler _SheetPivotTableAfterValueChangeEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.SheetPivotTableAfterValueChange"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public event Workbook_SheetPivotTableAfterValueChangeEventHandler SheetPivotTableAfterValueChangeEvent
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
		private event Workbook_SheetPivotTableBeforeAllocateChangesEventHandler _SheetPivotTableBeforeAllocateChangesEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.SheetPivotTableBeforeAllocateChanges"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public event Workbook_SheetPivotTableBeforeAllocateChangesEventHandler SheetPivotTableBeforeAllocateChangesEvent
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
		private event Workbook_SheetPivotTableBeforeCommitChangesEventHandler _SheetPivotTableBeforeCommitChangesEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.SheetPivotTableBeforeCommitChanges"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public event Workbook_SheetPivotTableBeforeCommitChangesEventHandler SheetPivotTableBeforeCommitChangesEvent
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
		private event Workbook_SheetPivotTableBeforeDiscardChangesEventHandler _SheetPivotTableBeforeDiscardChangesEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.SheetPivotTableBeforeDiscardChanges"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public event Workbook_SheetPivotTableBeforeDiscardChangesEventHandler SheetPivotTableBeforeDiscardChangesEvent
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
		private event Workbook_SheetPivotTableChangeSyncEventHandler _SheetPivotTableChangeSyncEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.SheetPivotTableChangeSync"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public event Workbook_SheetPivotTableChangeSyncEventHandler SheetPivotTableChangeSyncEvent
		{
			add
			{
				CreateEventBridge();
				_SheetPivotTableChangeSyncEvent += value;
			}
			remove
			{
				_SheetPivotTableChangeSyncEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Workbook_AfterSaveEventHandler _AfterSaveEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.AfterSave"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public event Workbook_AfterSaveEventHandler AfterSaveEvent
		{
			add
			{
				CreateEventBridge();
				_AfterSaveEvent += value;
			}
			remove
			{
				_AfterSaveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 14,15,16
		/// </summary>
		private event Workbook_NewChartEventHandler _NewChartEvent;

		/// <summary>
		/// SupportByVersion Excel 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbook.NewChart"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public event Workbook_NewChartEventHandler NewChartEvent
		{
			add
			{
				CreateEventBridge();
				_NewChartEvent += value;
			}
			remove
			{
				_NewChartEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 15, 16
		/// </summary>
		private event Workbook_SheetLensGalleryRenderCompleteEventHandler _SheetLensGalleryRenderCompleteEvent;

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.workbook.sheetlensgalleryrendercomplete"/> </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public event Workbook_SheetLensGalleryRenderCompleteEventHandler SheetLensGalleryRenderCompleteEvent
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
		private event Workbook_SheetTableUpdateEventHandler _SheetTableUpdateEvent;

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.workbook.sheettableupdate"/> </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public event Workbook_SheetTableUpdateEventHandler SheetTableUpdateEvent
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
		private event Workbook_ModelChangeEventHandler _ModelChangeEvent;

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.workbook.modelchange"/> </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public event Workbook_ModelChangeEventHandler ModelChangeEvent
		{
			add
			{
				CreateEventBridge();
				_ModelChangeEvent += value;
			}
			remove
			{
				_ModelChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Excel, 15, 16
		/// </summary>
		private event Workbook_SheetBeforeDeleteEventHandler _SheetBeforeDeleteEvent;

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.workbook.sheetbeforedelete"/> </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public event Workbook_SheetBeforeDeleteEventHandler SheetBeforeDeleteEvent
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, Events.WorkbookEvents_SinkHelper.Id);


			if(Events.WorkbookEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_workbookEvents_SinkHelper = new Events.WorkbookEvents_SinkHelper(this, _connectPoint);
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
			if( null != _workbookEvents_SinkHelper)
			{
				_workbookEvents_SinkHelper.Dispose();
				_workbookEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}

