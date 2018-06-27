using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CoreServices;
using NetOffice.CollectionsGeneric;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// CoClass Application
    /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194565.aspx </remarks>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass), ComProgId("Excel.Application"), ModuleProvider(typeof(ModulesLegacy.ApplicationModule))]
    [ComEventContract(typeof(NetOffice.ExcelApi.EventContracts.AppEvents))]
    [HasInteropCompatibilityClass(typeof(ApplicationClass))]
    public class Application : NetOffice.ExcelApi.Behind._Application, NetOffice.ExcelApi.Application, IAutomaticQuit, IApplicationVersionProvider
    {
        #pragma warning disable

        #region Fields

        private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
        private string _activeSinkId;
        private static Type _type;
        private EventContracts.AppEvents_SinkHelper _appEvents_SinkHelper;

        private bool _versionRequested;
        private object _cachedVersion;
        private object _chachedVersionLock = new object();

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
                    _contractType = typeof(NetOffice.ExcelApi.Application);
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
                    _type = typeof(Application);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public Application() : base()
        {

        }

        /// <summary>
        /// Creates a new instance of the class
        /// <param name="factory">factory core instead of default core</param>
        /// <param name="tryProxyServiceFirst">try to get a running application first before create a new application</param>
        /// </summary>
        public Application(Core factory = null, bool tryProxyServiceFirst = false) : base()
        {
            object proxy = null;
            if (tryProxyServiceFirst)
            {
                proxy = ProxyService.GetActiveInstance("Excel", "Application", false);
                if (null != proxy)
                {
                    CreateFromProxy(proxy, true);
                    FromProxyService = true;
                }
            }

            if(null == proxy)
            {
                CreateFromProgId("Excel.Application", true);
            }

            Factory = null != factory ? factory : Core.Default;
            TryRequestVersion();
            RegisterAsApplicationVersionProvider();
            OnCreate();
            _isInitialized = true;
        }

        #endregion

        #region ICloneable<Application>

        /// <summary>
        /// Creates a new Application instance that is a copy of the current instance
        /// </summary>
        /// <returns>A new Application that is a copy of this instance</returns>
        /// <exception cref="NetOffice.Exceptions.CloneException">An unexpected error occured. See inner exception(s) for details.</exception>
        public new virtual ExcelApi.Application DeepCopy()
        {
            return base.Clone() as ExcelApi.Application;
        }

        #endregion

        #region ICOMObjectProxyService

        /// <summary>
        /// Instance is created from an already running application
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public virtual bool FromProxyService { get; private set; }

        #endregion

        #region IAutomaticQuit

        /// <summary>
        /// Determines Quit method want be called while disposing if NetOffice.Settings.EnableAutomaticQuit is true.
        /// Default is true when instance has no parent object and its not a cloned instance, otherwise false.
        /// </summary>
        bool IAutomaticQuit.Enabled
        {

            get
            {
                return _callQuitInDispose;
            }
            set
            {
                _callQuitInDispose = value;
            }
        }

        #endregion

        #region IApplicationVersionProvider

        string IApplicationVersionProvider.Name
        {
            get
            {
                return "Microsoft Excel";
            }
        }

        string IApplicationVersionProvider.ComponentName
        {
            get
            {
                return "NetOffice.ExcelApi";
            }
        }

        /// <summary>
        /// Request version information on demand and cache to call the remote server only 1x times
        /// </summary>
        object IApplicationVersionProvider.Version
        {
            get
            {
                lock (_chachedVersionLock)
                {
                    if (null == _cachedVersion)
                    {
                        _cachedVersion = TryVersionPropertyGet();
                    }
                }
                return _cachedVersion;
            }
        }

        bool IApplicationVersionProvider.VersionRequested
        {
            get
            {
                return _versionRequested;
            }
        }

        void IApplicationVersionProvider.TryRequestVersion()
        {
            _cachedVersion = TryVersionPropertyGet();
        }

        /// <summary>
        /// Try get version information without fail
        /// </summary>
        /// <returns></returns>
        private object TryVersionPropertyGet()
        {
            try
            {
                if (null != _proxyShare)
                    return Invoker.PropertyGet(this, "Version");
                else
                    return null;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (null != _proxyShare)
                    _versionRequested = true;
            }
        }

        #endregion

        #region Static CoClass Methods

        /// <summary>
        /// Returns all running Excel.Application instances from the environment/system
        /// </summary>
        /// <returns>Excel.Application sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances()
        {
            return ProxyService.GetActiveInstances<Application>("Excel", "Application");
        }

        /// <summary>
        /// Returns all running Excel.Application instances from the environment/system that passed a predicate filter
        /// </summary>
        /// <param name="predicate">filter predicate</param>
        /// <returns>Excel.Application sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances(Func<Application, bool> predicate)
        {
            return ProxyService.GetActiveInstances<Application>("Excel", "Application", predicate);
        }

        /// <summary>
        /// Returns the count of running Excel.Application instances that passed a predicate filter
        /// </summary>
        /// <returns>count of running application</returns>
        public static int GetActiveInstancesCount()
        {
            using (var sequence = ProxyService.GetActiveInstances<Application>("Excel", "Application"))
            {
                int result = sequence.Count;
                sequence.Dispose();
                return result;
            }            
        }

        /// <summary>
        /// Returns the count of running Excel.Application instances that passed a predicate filter
        /// </summary>
        /// <param name="predicate">filter predicate</param>
        /// <returns>count of running application</returns>
        public static int GetActiveInstancesCount(Func<Application, bool> predicate)
        {
            using (var sequence = ProxyService.GetActiveInstances<Application>("Excel", "Application", predicate))
            {
                int result = sequence.Count;
                sequence.Dispose();
                return result;
            }
        }

        /// <summary>
        /// Returns a running Excel.Application instance from the environment/system
        /// </summary>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>Excel.Application instance or null(Nothing in Visual Basic)</returns>
        /// <exception cref="ArgumentOutOfRangeException">occurs if no instance match and throwExceptionIfNotFound is set</exception>
        public static Application GetActiveInstance(bool throwExceptionIfNotFound = false)
        {
            return ProxyService.GetActiveInstance<Application>("Excel", "Application", throwExceptionIfNotFound);
        }

        /// <summary>
        /// Returns first running Excel.Application instance from the environment/system that passed a predicate filter
        /// </summary>
        /// <param name="predicate">filter predicate</param>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>Excel.Application instance or null(Nothing in Visual Basic)</returns>
        /// <exception cref="ArgumentOutOfRangeException">occurs if no instance match and throwExceptionIfNotFound is set</exception>
        public static Application GetActiveInstance(Func<Application, bool> predicate, bool throwExceptionIfNotFound = false)
        {
            return ProxyService.GetActiveInstance<Application>("Excel", "Application", predicate, throwExceptionIfNotFound);
        }

        #endregion

        #region Events

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Application_NewWorkbookEventHandler _NewWorkbookEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837373.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_NewWorkbookEventHandler NewWorkbookEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Application_SheetSelectionChangeEventHandler _SheetSelectionChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839035.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_SheetSelectionChangeEventHandler SheetSelectionChangeEvent
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
        private event Application_SheetBeforeDoubleClickEventHandler _SheetBeforeDoubleClickEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836225.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClickEvent
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
        private event Application_SheetBeforeRightClickEventHandler _SheetBeforeRightClickEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840532.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_SheetBeforeRightClickEventHandler SheetBeforeRightClickEvent
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
        private event Application_SheetActivateEventHandler _SheetActivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193288.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_SheetActivateEventHandler SheetActivateEvent
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
        private event Application_SheetDeactivateEventHandler _SheetDeactivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823120.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_SheetDeactivateEventHandler SheetDeactivateEvent
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
        private event Application_SheetCalculateEventHandler _SheetCalculateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835607.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_SheetCalculateEventHandler SheetCalculateEvent
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
        private event Application_SheetChangeEventHandler _SheetChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193591.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_SheetChangeEventHandler SheetChangeEvent
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
        private event Application_WorkbookOpenEventHandler _WorkbookOpenEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196583.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_WorkbookOpenEventHandler WorkbookOpenEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Application_WorkbookActivateEventHandler _WorkbookActivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837347.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_WorkbookActivateEventHandler WorkbookActivateEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Application_WorkbookDeactivateEventHandler _WorkbookDeactivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193560.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_WorkbookDeactivateEventHandler WorkbookDeactivateEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Application_WorkbookBeforeCloseEventHandler _WorkbookBeforeCloseEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836770.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_WorkbookBeforeCloseEventHandler WorkbookBeforeCloseEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Application_WorkbookBeforeSaveEventHandler _WorkbookBeforeSaveEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840422.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_WorkbookBeforeSaveEventHandler WorkbookBeforeSaveEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Application_WorkbookBeforePrintEventHandler _WorkbookBeforePrintEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195507.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_WorkbookBeforePrintEventHandler WorkbookBeforePrintEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Application_WorkbookNewSheetEventHandler _WorkbookNewSheetEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198367.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_WorkbookNewSheetEventHandler WorkbookNewSheetEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Application_WorkbookAddinInstallEventHandler _WorkbookAddinInstallEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836206.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_WorkbookAddinInstallEventHandler WorkbookAddinInstallEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Application_WorkbookAddinUninstallEventHandler _WorkbookAddinUninstallEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835570.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_WorkbookAddinUninstallEventHandler WorkbookAddinUninstallEvent
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
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Application_WindowResizeEventHandler _WindowResizeEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836166.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_WindowResizeEventHandler WindowResizeEvent
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
        private event Application_WindowActivateEventHandler _WindowActivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821328.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_WindowActivateEventHandler WindowActivateEvent
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
        private event Application_WindowDeactivateEventHandler _WindowDeactivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822473.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_WindowDeactivateEventHandler WindowDeactivateEvent
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
        private event Application_SheetFollowHyperlinkEventHandler _SheetFollowHyperlinkEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821956.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Application_SheetFollowHyperlinkEventHandler SheetFollowHyperlinkEvent
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
        private event Application_SheetPivotTableUpdateEventHandler _SheetPivotTableUpdateEvent;

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840950.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual event Application_SheetPivotTableUpdateEventHandler SheetPivotTableUpdateEvent
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
        private event Application_WorkbookPivotTableCloseConnectionEventHandler _WorkbookPivotTableCloseConnectionEvent;

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198029.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual event Application_WorkbookPivotTableCloseConnectionEventHandler WorkbookPivotTableCloseConnectionEvent
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
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Application_WorkbookPivotTableOpenConnectionEventHandler _WorkbookPivotTableOpenConnectionEvent;

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821547.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual event Application_WorkbookPivotTableOpenConnectionEventHandler WorkbookPivotTableOpenConnectionEvent
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
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        private event Application_WorkbookSyncEventHandler _WorkbookSyncEvent;

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839042.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual event Application_WorkbookSyncEventHandler WorkbookSyncEvent
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
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        private event Application_WorkbookBeforeXmlImportEventHandler _WorkbookBeforeXmlImportEvent;

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196324.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual event Application_WorkbookBeforeXmlImportEventHandler WorkbookBeforeXmlImportEvent
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
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        private event Application_WorkbookAfterXmlImportEventHandler _WorkbookAfterXmlImportEvent;

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837416.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual event Application_WorkbookAfterXmlImportEventHandler WorkbookAfterXmlImportEvent
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
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        private event Application_WorkbookBeforeXmlExportEventHandler _WorkbookBeforeXmlExportEvent;

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195824.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual event Application_WorkbookBeforeXmlExportEventHandler WorkbookBeforeXmlExportEvent
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
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        private event Application_WorkbookAfterXmlExportEventHandler _WorkbookAfterXmlExportEvent;

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836803.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual event Application_WorkbookAfterXmlExportEventHandler WorkbookAfterXmlExportEvent
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
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        private event Application_WorkbookRowsetCompleteEventHandler _WorkbookRowsetCompleteEvent;

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839165.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual event Application_WorkbookRowsetCompleteEventHandler WorkbookRowsetCompleteEvent
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
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        private event Application_AfterCalculateEventHandler _AfterCalculateEvent;

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840621.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual event Application_AfterCalculateEventHandler AfterCalculateEvent
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
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Application_SheetPivotTableAfterValueChangeEventHandler _SheetPivotTableAfterValueChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193316.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Application_SheetPivotTableAfterValueChangeEventHandler SheetPivotTableAfterValueChangeEvent
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
        private event Application_SheetPivotTableBeforeAllocateChangesEventHandler _SheetPivotTableBeforeAllocateChangesEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838226.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Application_SheetPivotTableBeforeAllocateChangesEventHandler SheetPivotTableBeforeAllocateChangesEvent
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
        private event Application_SheetPivotTableBeforeCommitChangesEventHandler _SheetPivotTableBeforeCommitChangesEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838379.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Application_SheetPivotTableBeforeCommitChangesEventHandler SheetPivotTableBeforeCommitChangesEvent
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
        private event Application_SheetPivotTableBeforeDiscardChangesEventHandler _SheetPivotTableBeforeDiscardChangesEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835217.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Application_SheetPivotTableBeforeDiscardChangesEventHandler SheetPivotTableBeforeDiscardChangesEvent
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
        private event Application_ProtectedViewWindowOpenEventHandler _ProtectedViewWindowOpenEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194431.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Application_ProtectedViewWindowOpenEventHandler ProtectedViewWindowOpenEvent
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
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Application_ProtectedViewWindowBeforeEditEventHandler _ProtectedViewWindowBeforeEditEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838239.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Application_ProtectedViewWindowBeforeEditEventHandler ProtectedViewWindowBeforeEditEvent
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
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Application_ProtectedViewWindowBeforeCloseEventHandler _ProtectedViewWindowBeforeCloseEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821579.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Application_ProtectedViewWindowBeforeCloseEventHandler ProtectedViewWindowBeforeCloseEvent
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
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Application_ProtectedViewWindowResizeEventHandler _ProtectedViewWindowResizeEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836848.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Application_ProtectedViewWindowResizeEventHandler ProtectedViewWindowResizeEvent
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
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Application_ProtectedViewWindowActivateEventHandler _ProtectedViewWindowActivateEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195451.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Application_ProtectedViewWindowActivateEventHandler ProtectedViewWindowActivateEvent
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
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Application_ProtectedViewWindowDeactivateEventHandler _ProtectedViewWindowDeactivateEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196820.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Application_ProtectedViewWindowDeactivateEventHandler ProtectedViewWindowDeactivateEvent
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
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Application_WorkbookAfterSaveEventHandler _WorkbookAfterSaveEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198184.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Application_WorkbookAfterSaveEventHandler WorkbookAfterSaveEvent
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
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Application_WorkbookNewChartEventHandler _WorkbookNewChartEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834985.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Application_WorkbookNewChartEventHandler WorkbookNewChartEvent
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
        /// SupportByVersion Excel 15, 16
        /// </summary>
        private event Application_SheetLensGalleryRenderCompleteEventHandler _SheetLensGalleryRenderCompleteEvent;

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227506.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        public virtual event Application_SheetLensGalleryRenderCompleteEventHandler SheetLensGalleryRenderCompleteEvent
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
        private event Application_SheetTableUpdateEventHandler _SheetTableUpdateEvent;

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229805.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        public virtual event Application_SheetTableUpdateEventHandler SheetTableUpdateEvent
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
        private event Application_WorkbookModelChangeEventHandler _WorkbookModelChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229611.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        public virtual event Application_WorkbookModelChangeEventHandler WorkbookModelChangeEvent
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
        /// SupportByVersion Excel 15, 16
        /// </summary>
        private event Application_SheetBeforeDeleteEventHandler _SheetBeforeDeleteEvent;

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/dn448391.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        public virtual event Application_SheetBeforeDeleteEventHandler SheetBeforeDeleteEvent
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
                _activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, EventContracts.AppEvents_SinkHelper.Id);

            if (EventContracts.AppEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
            {
                _appEvents_SinkHelper = new EventContracts.AppEvents_SinkHelper(this, _connectPoint);
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
            if (null != _appEvents_SinkHelper)
            {
                _appEvents_SinkHelper.Dispose();
                _appEvents_SinkHelper = null;
            }

            _connectPoint = null;
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Called from ctor or ICOMObjectInitialize at last as an inherited class service
        /// </summary>
        protected override void OnCreate()
        {
            if (null == ParentObject)
            {
                _callQuitInDispose = true;
                ModulesLegacy.ApplicationModule.Instance = this;
            }
            base.OnCreate();
        }

        /// <summary>
        /// NetOffice method: dispose instance and all child instances
        /// </summary>
        /// <param name="disposeEventBinding">dispose event exported proxies with one or more event recipients</param>
        [Category("NetOffice"), CoreOverridden]
        public override void Dispose(bool disposeEventBinding)
        {
            if (this.Equals(ModulesLegacy.ApplicationModule.Instance))
                ModulesLegacy.ApplicationModule.Instance = null;
            base.Dispose(disposeEventBinding);
        }

        /// <summary>
        /// NetOffice method: dispose instance and all child instances
        /// </summary>
        [Category("NetOffice"), CoreOverridden]
        public override void Dispose()
        {
            if (this.Equals(ModulesLegacy.ApplicationModule.Instance))
                ModulesLegacy.ApplicationModule.Instance = null;
            base.Dispose();
        }

        #endregion

        #pragma warning restore
    }
}
