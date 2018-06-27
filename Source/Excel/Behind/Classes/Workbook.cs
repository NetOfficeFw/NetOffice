using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// CoClass Workbook
    /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835568.aspx </remarks>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.ExcelApi.EventContracts.WorkbookEvents))]
    public class Workbook : NetOffice.ExcelApi.Behind._Workbook, NetOffice.ExcelApi.Workbook
    {
        #pragma warning disable

        #region Fields

        private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
        private string _activeSinkId;
        private static Type _type;
        private EventContracts.WorkbookEvents_SinkHelper _workbookEvents_SinkHelper;

        #endregion

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
                    _contractType = typeof(NetOffice.ExcelApi.Workbook);
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

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public Workbook() : base()
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
            return ProxyService.GetActiveInstances<Application>("Excel", "Workbook");
        }

        /// <summary>
        /// Returns a running Excel.Workbook instance from the environment/system
        /// </summary>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>Excel.Workbook instance or null</returns>
        public static Application GetActiveInstance(bool throwExceptionIfNotFound = false)
        {
            return ProxyService.GetActiveInstance<Application>("Excel", "Workbook", throwExceptionIfNotFound);
        }

        #endregion

        #region Events

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_OpenEventHandler _OpenEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196215.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_OpenEventHandler OpenEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_ActivateEventHandler _ActivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823078.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_ActivateEventHandler ActivateEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_DeactivateEventHandler _DeactivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822521.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_DeactivateEventHandler DeactivateEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_BeforeCloseEventHandler _BeforeCloseEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194765.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_BeforeCloseEventHandler BeforeCloseEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_BeforeSaveEventHandler _BeforeSaveEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840057.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_BeforeSaveEventHandler BeforeSaveEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_BeforePrintEventHandler _BeforePrintEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195836.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_BeforePrintEventHandler BeforePrintEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_NewSheetEventHandler _NewSheetEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821246.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_NewSheetEventHandler NewSheetEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_AddinInstallEventHandler _AddinInstallEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822158.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_AddinInstallEventHandler AddinInstallEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_AddinUninstallEventHandler _AddinUninstallEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840207.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_AddinUninstallEventHandler AddinUninstallEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_WindowResizeEventHandler _WindowResizeEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822648.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_WindowResizeEventHandler WindowResizeEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_WindowActivateEventHandler _WindowActivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840441.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_WindowActivateEventHandler WindowActivateEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_WindowDeactivateEventHandler _WindowDeactivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839676.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_WindowDeactivateEventHandler WindowDeactivateEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_SheetSelectionChangeEventHandler _SheetSelectionChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837368.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_SheetSelectionChangeEventHandler SheetSelectionChangeEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_SheetBeforeDoubleClickEventHandler _SheetBeforeDoubleClickEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822360.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClickEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_SheetBeforeRightClickEventHandler _SheetBeforeRightClickEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839675.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_SheetBeforeRightClickEventHandler SheetBeforeRightClickEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_SheetActivateEventHandler _SheetActivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195710.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_SheetActivateEventHandler SheetActivateEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_SheetDeactivateEventHandler _SheetDeactivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838589.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_SheetDeactivateEventHandler SheetDeactivateEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_SheetCalculateEventHandler _SheetCalculateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193282.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_SheetCalculateEventHandler SheetCalculateEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_SheetChangeEventHandler _SheetChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196611.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_SheetChangeEventHandler SheetChangeEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_SheetFollowHyperlinkEventHandler _SheetFollowHyperlinkEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838573.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_SheetFollowHyperlinkEventHandler SheetFollowHyperlinkEvent
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
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_SheetPivotTableUpdateEventHandler _SheetPivotTableUpdateEvent;

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193598.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_SheetPivotTableUpdateEventHandler SheetPivotTableUpdateEvent
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
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_PivotTableCloseConnectionEventHandler _PivotTableCloseConnectionEvent;

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231875.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_PivotTableCloseConnectionEventHandler PivotTableCloseConnectionEvent
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
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_PivotTableOpenConnectionEventHandler _PivotTableOpenConnectionEvent;

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231367.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual event Workbook_PivotTableOpenConnectionEventHandler PivotTableOpenConnectionEvent
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
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_SyncEventHandler _SyncEvent;

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231646.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual event Workbook_SyncEventHandler SyncEvent
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
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_BeforeXmlImportEventHandler _BeforeXmlImportEvent;

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837099.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual event Workbook_BeforeXmlImportEventHandler BeforeXmlImportEvent
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
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_AfterXmlImportEventHandler _AfterXmlImportEvent;

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838071.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual event Workbook_AfterXmlImportEventHandler AfterXmlImportEvent
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
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_BeforeXmlExportEventHandler _BeforeXmlExportEvent;

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840659.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual event Workbook_BeforeXmlExportEventHandler BeforeXmlExportEvent
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
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        private event Workbook_AfterXmlExportEventHandler _AfterXmlExportEvent;

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841244.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual event Workbook_AfterXmlExportEventHandler AfterXmlExportEvent
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
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        private event Workbook_RowsetCompleteEventHandler _RowsetCompleteEvent;

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193275.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual event Workbook_RowsetCompleteEventHandler RowsetCompleteEvent
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
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Workbook_SheetPivotTableAfterValueChangeEventHandler _SheetPivotTableAfterValueChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834987.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Workbook_SheetPivotTableAfterValueChangeEventHandler SheetPivotTableAfterValueChangeEvent
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
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Workbook_SheetPivotTableBeforeAllocateChangesEventHandler _SheetPivotTableBeforeAllocateChangesEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196066.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Workbook_SheetPivotTableBeforeAllocateChangesEventHandler SheetPivotTableBeforeAllocateChangesEvent
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
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Workbook_SheetPivotTableBeforeCommitChangesEventHandler _SheetPivotTableBeforeCommitChangesEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834653.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Workbook_SheetPivotTableBeforeCommitChangesEventHandler SheetPivotTableBeforeCommitChangesEvent
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
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Workbook_SheetPivotTableBeforeDiscardChangesEventHandler _SheetPivotTableBeforeDiscardChangesEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840414.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Workbook_SheetPivotTableBeforeDiscardChangesEventHandler SheetPivotTableBeforeDiscardChangesEvent
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
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Workbook_SheetPivotTableChangeSyncEventHandler _SheetPivotTableChangeSyncEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838762.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Workbook_SheetPivotTableChangeSyncEventHandler SheetPivotTableChangeSyncEvent
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
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Workbook_AfterSaveEventHandler _AfterSaveEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836466.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Workbook_AfterSaveEventHandler AfterSaveEvent
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
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Workbook_NewChartEventHandler _NewChartEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823186.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Workbook_NewChartEventHandler NewChartEvent
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
        /// SupportByVersion Excel 15, 16
        /// </summary>
        private event Workbook_SheetLensGalleryRenderCompleteEventHandler _SheetLensGalleryRenderCompleteEvent;

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230636.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        public virtual event Workbook_SheetLensGalleryRenderCompleteEventHandler SheetLensGalleryRenderCompleteEvent
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
        /// SupportByVersion Excel 15, 16
        /// </summary>
        private event Workbook_SheetTableUpdateEventHandler _SheetTableUpdateEvent;

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229546.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        public virtual event Workbook_SheetTableUpdateEventHandler SheetTableUpdateEvent
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
        /// SupportByVersion Excel 15, 16
        /// </summary>
        private event Workbook_ModelChangeEventHandler _ModelChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232031.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        public virtual event Workbook_ModelChangeEventHandler ModelChangeEvent
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
        /// SupportByVersion Excel 15, 16
        /// </summary>
        private event Workbook_SheetBeforeDeleteEventHandler _SheetBeforeDeleteEvent;

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/dn448396.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        public virtual event Workbook_SheetBeforeDeleteEventHandler SheetBeforeDeleteEvent
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
        public virtual void CreateEventBridge()
        {
            if (false == Factory.Settings.EnableEvents)
                return;

            if (null != _connectPoint)
                return;

            if (null == _activeSinkId)
                _activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, EventContracts.WorkbookEvents_SinkHelper.Id);


            if (EventContracts.WorkbookEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
            {
                _workbookEvents_SinkHelper = new EventContracts.WorkbookEvents_SinkHelper(this, _connectPoint);
                return;
            }
        }

        /// <summary>
        /// The instance use currently an event listener
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool EventBridgeInitialized
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
        public virtual bool HasEventRecipients()
        {
            return NetOffice.Events.CoClassEventReflector.HasEventRecipients(this, LateBindingApiWrapperType);
        }

        /// <summary>
        /// Instance has one or more event recipients
        /// </summary>
        /// <param name="eventName">name of the event</param>
        /// <returns></returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool HasEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.HasEventRecipients(this, LateBindingApiWrapperType, eventName);
        }

        /// <summary>
        /// Target methods from its actual event recipients
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Delegate[] GetEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.GetEventRecipients(this, LateBindingApiWrapperType, eventName);
        }

        /// <summary>
        /// Returns the current count of event recipients
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual int GetCountOfEventRecipients(string eventName)
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
        public virtual int RaiseCustomEvent(string eventName, ref object[] paramsArray)
        {
            return NetOffice.Events.CoClassEventReflector.RaiseCustomEvent(this, LateBindingApiWrapperType, eventName, ref paramsArray);
        }
        /// <summary>
        /// Stop listening events for the instance
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void DisposeEventBridge()
        {
            if (null != _workbookEvents_SinkHelper)
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
