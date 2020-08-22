using System;
using System.ComponentModel;
using System.Globalization;
using NetOffice.Filtering;

namespace NetOffice
{
    /// <summary>
    /// Core Settings
    /// </summary>
    public class Settings : INotifyPropertyChanged, IEquatable<Settings>
    {
        #region Constants

        // Default Thread Culture
        private const string _DefaultCulture = "en-US";

        #endregion

        #region Fields

        private CultureInfo _cultureInfo;

        private bool _eventsEnabled = true;
        private RetryMessageFilter _messageFilter;
        private bool _enableAutomaticQuit = false;
        private bool _enableAdHocLoading = true;
        private bool _enableDeepLoading = true;
        private bool _enableMoreDebugOutput = false;
        private bool _enableEventDebugOutput = false;
        private bool _enableSafeMode = false;
        private bool _enableDynamicObjects = true;
        private bool _enableDynamicEventArguments;
        private bool _enableKnownReferenceInspection;
        private bool _enableAutoDisposeEventArguments;

        private bool _enableProxyManagement = true;
        private CacheOptions _cacheOptions = CacheOptions.KeepExistingCacheAlive;
        private bool _enableOperatorOverloads = true;
        private string _exceptionDefaultMessage = "See inner exception(s) for details.";
        private string _exceptionDiagnosticsMessage = "Failed to proceed {CallType} on {CallInstance}=>{Name}.";
        private ExceptionMessageHandling _exceptionMessageBehavior = ExceptionMessageHandling.Diagnostics;
        private bool _loadAssembliesUnsafe = true;
        private PerformanceTrace _performanceTrace;
        private static Settings _default;
        private static object _defaultLock = new object();

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public Settings()
        {
            _messageFilter = new RetryMessageFilter(OnPropertyChanged);
            _performanceTrace = new PerformanceTrace(OnPropertyChanged);
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="settings">settings to load from</param>
        public Settings(Settings settings)
        {
            _messageFilter = new RetryMessageFilter(OnPropertyChanged);
            _performanceTrace = new PerformanceTrace(OnPropertyChanged);
            LoadFrom(settings);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Shared Default Settings
        /// </summary>
        public static Settings Default
        {
            get
            {
                lock (_defaultLock)
                {
                    if (null == _default)
                        _default = new Settings();
                }
                return _default;
            }
        }

        /// <summary>
        /// Performance tracer to see how long its need to call and return all or specific actions
        /// </summary>
        [Category("Settings"), Description("Trace system to measure performance.")]    
        public PerformanceTrace PerformanceTrace
        {
            get
            {
                return _performanceTrace;
            }
        }

        /// <summary>
        /// Enable the NetOffice COM proxy management. true by default
        /// </summary>
        [Category("Settings"), Description("Enable the COM proxy management in a parent child relation."), DefaultValue(true)]
        public bool EnableProxyManagement
        {
            get
            {
                return _enableProxyManagement;
            }
            set
            {
                if (value != _enableProxyManagement)
                { 
                    _enableProxyManagement = value;
                    OnPropertyChanged("EnableProxyManagement");
                }
            }
        }
        
        /// <summary>
        /// Analyze also known reference proxies to see proxy is may inherited type, false by default
        /// </summary>
        [Category("Settings"), Description("Analyze also known reference proxies to see proxy is may inherited type."), DefaultValue(false)]
        public bool EnableKnownReferenceInspection
        {
            get
            {
                return _enableKnownReferenceInspection;
            }
            set
            {
                if (value != _enableKnownReferenceInspection)
                {
                    _enableKnownReferenceInspection = value;
                    OnPropertyChanged("EnableKnownReferenceInspection");
                }
            }
        }

        /// <summary>
        /// Dispose event arguments automatically, false by default
        /// </summary>
        [Category("Settings"), Description("Dispose event arguments automatically."), DefaultValue(false)]
        public bool EnableAutoDisposeEventArguments
        {
            get
            {
                return _enableAutoDisposeEventArguments;
            }
            set
            {
                if (value != _enableAutoDisposeEventArguments)
                {
                    _enableAutoDisposeEventArguments = value;
                    OnPropertyChanged("EnableAutoDisposeEventArguments");
                }
            }
        }

        /// <summary>
        /// NetOffice wrap all thrown exceptions from Office applications in a COMException.
        /// This option can be used to set the top level exception message or copy the inner message to top.
        /// </summary>
        [Category("Settings"), Description("Determine exception message behavior."), DefaultValue(typeof(ExceptionMessageHandling), "Diagnostics")]
        public ExceptionMessageHandling ExceptionMessageBehavior
        {
            get
            {
                return _exceptionMessageBehavior;
            }
            set
            {
                if (value != _exceptionMessageBehavior)
                { 
                    _exceptionMessageBehavior = value;
                    OnPropertyChanged("ExceptionMessageBehavior");
                }
            }
        }

        /// <summary>
        /// NetOffice wrap all thrown exceptions from Office applications in a COMException.
        /// This is the default message for the top level exception when ExceptionMessageBehavior is Default.
        /// </summary>
        [Category("Settings"), Description("Top Level Exception default message text."), DefaultValue("See inner exception(s) for details.")]
        public string ExceptionDefaultMessage
        {
            get
            {
                return _exceptionDefaultMessage;
            }
            set
            {
                if (value != _exceptionDefaultMessage)
                {
                    _exceptionDefaultMessage = value;
                    OnPropertyChanged("ExceptionDefaultMessage");
                }
            }
        }

        /// <summary>
        /// NetOffice wrap all thrown exceptions from Office applications in a COMException.
        /// This is the default message for the top level exception when ExceptionMessageBehavior is Diagnostics.
        /// See ExceptionMessageHandling.Diagnostics for further information.
        /// </summary>
        [Category("Settings"), Description("Top Level  exception diagnostics message text."), DefaultValue("Failed to proceed {CallType} on {CallInstance}=>{Name}.")]
        public string ExceptionDiagnosticsMessage
        {
            get
            {
                return _exceptionDiagnosticsMessage;
            }
            set
            {
                if (value != _exceptionDiagnosticsMessage)
                {
                    _exceptionDiagnosticsMessage = value;
                    OnPropertyChanged("ExceptionDiagnosticsMessage");
                }
            }
        }

        /// <summary>
        /// Used Thread Culture given in the invoke calls. en-US by default
        /// </summary>
        [Category("Settings"), Description("Given thread culture in remote server calls."), DefaultValue(typeof(CultureInfo), "en-us")]
        public CultureInfo ThreadCulture
        {
            get
            {
                try
                {
                    if (null == _cultureInfo)
                        _cultureInfo = CultureInfo.GetCultureInfo(_DefaultCulture);
                }
                catch (Exception throwedException)
                {
                    throw (new NetOfficeException(string.Format("GetCultureInfo {0} failed.", _DefaultCulture), throwedException));
                }
                finally
                {
                    if (null == _cultureInfo)
                        throw (new NetOfficeException(string.Format("GetCultureInfo {0} failed.", _DefaultCulture)));
                }

                return _cultureInfo;
            }
            set
            {
                if (null == value)
                    throw new ArgumentNullException(nameof(value), "ThreadCulture must have a value");

                if (value != _cultureInfo)
                {
                    _cultureInfo = value;
                    OnPropertyChanged("ThreadCulture");
                }
            }
        }

        /// <summary>
        /// Get or set the Event support. true by default
        /// </summary>
        [Category("Settings"), Description("Enable or disable event subsystem."), DefaultValue(true)]
        public bool EnableEvents
        {
            get
            {
                return _eventsEnabled;
            }
            set
            {
                if (value != _eventsEnabled)
                {
                    _eventsEnabled = value;
                    OnPropertyChanged("EnableEvents");
                }
            }
        }

        /// <summary>
        /// A predefined implementation of IMessageFilter
        /// </summary>
        [Category("Settings"), Description("Predefined Message Filter Support")]
        public RetryMessageFilter MessageFilter
        {
            get
            {
                return _messageFilter;
            }
        }

        /// <summary>
        /// Get or set the Quit method for an application object was automatically called while Dispose. false by default
        /// </summary>
        [Category("Settings"), Description("Call Quit in dispose automatically if the instance support a Quit method."), DefaultValue(false)]
        public bool EnableAutomaticQuit
        {
            get
            {
                return _enableAutomaticQuit;
            }
            set
            {
                if (value != _enableAutomaticQuit)
                {
                    _enableAutomaticQuit = value;
                    OnPropertyChanged("EnableAutomaticQuit");
                }
            }
        }

        /// <summary>
        /// Get or set the core api checks at runtime the target method or property is supported in current version. if it doesn't a EntityNotSupportedException is thrown. false by default
        /// </summary>
        [Category("Settings"), Description("Check method or property is supported before call them and throw an EntityNotSupportedException if unable to find."), DefaultValue(false)]
        public bool EnableSafeMode
        {
            get
            {
                return _enableSafeMode;
            }
            set
            {
                if (value != _enableSafeMode)
                {
                    _enableSafeMode = value;
                    OnPropertyChanged("EnableSafeMode");
                }
            }
        }

        /// <summary>
        /// Get or set Core.Initialize() try to load unloaded dependent assemblies to fetch type information. true by default
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [Category("Settings"), Description("Load assemblies from file system while initialize."), DefaultValue(true)]
        public bool EnableAdHocLoading
        {
            get
            {
                return _enableAdHocLoading;
            }
            set
            {
                if (value != _enableAdHocLoading)
                {
                    _enableAdHocLoading = value;
                    OnPropertyChanged("EnableAdHocLoading");
                }
            }
        }

        /// <summary>
        /// Get or set the Initialize method perform a deep level analyzing(may cause security issues). true by default
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [Category("Settings"), Description("Analyze the current AppDomain in detail to find necessary assemblies."), DefaultValue(true)]
        public bool EnableDeepLoading
        {
            get
            {
                return _enableDeepLoading;
            }
            set
            {
                if (value != _enableDeepLoading)
                {
                    _enableDeepLoading = value;
                    OnPropertyChanged("EnableDeepLoading");
                }
            }
        }

        /// <summary>
        /// Write extended diagnostics to console
        /// </summary>
        [Category("Settings"), Description("Write extended diagnostics to console."), DefaultValue(false)]
        public bool EnableMoreDebugOutput
        {
            get
            {
                return _enableMoreDebugOutput;
            }
            set
            {
                if (value != _enableMoreDebugOutput)
                {
                    _enableMoreDebugOutput = value;
                    OnPropertyChanged("EnableMoreDebugOutput");
                }
            }
        }

        /// <summary>
        /// Get or set NetOffice logs essential system steps for event operations in the DebugConsole(if enabled). false by default
        /// </summary>
        [Category("Settings"), Description("Event related debug messages want be shown in console."), DefaultValue(false)]
        public bool EnableEventDebugOutput
        {
            get
            {
                return _enableEventDebugOutput;
            }
            set
            {
                if (value != _enableEventDebugOutput)
                {
                    _enableEventDebugOutput = value;
                    OnPropertyChanged("EnableEventDebugOutput");
                }
            }
        }

        /// <summary>
        /// Get or set Core.Initialize() try to load non loaded dependent assemblies to fetch type information. KeepExistingCacheAlive by default
        /// </summary>
        [Category("Settings"), Description("Re-use or skip existing information while initialize."), DefaultValue(typeof(CacheOptions), "KeepExistingCacheAlive")]
        public CacheOptions CacheOptions
        {
            get
            {
                return _cacheOptions;
            }
            set
            {
                if (value != _cacheOptions)
                {
                    _cacheOptions = value;
                    OnPropertyChanged("CacheOptions");
                }
            }
        }

        /// <summary>
        /// Get or set NetOffice should create custom overloads for the "==" and "!=" operators for semantic comparisons. true by default
        /// </summary>
        [Category("Settings"), Description("Redirect equal operations like '==' or '!=' for proxy wrapping objects to the com server to determine 2 instances are equal."), DefaultValue(true)]
        public bool EnableOperatorOverloads
        {
            get
            {
                return _enableOperatorOverloads;
            }
            set
            {
                if (value != _enableOperatorOverloads)
                {
                    _enableOperatorOverloads = value;
                    OnPropertyChanged("EnableOperatorOverloads");
                }
            }
        }

        /// <summary>
        /// Get or set NetOffice try load dependent assemblies unsafe(System.Reflection.Assembly.UnsafeLoadFrom). true by default
        /// </summary>
        [Category("Settings"), Description("Load assemblies unsafe and bypass some security checks."), DefaultValue(true)]
        public bool LoadAssembliesUnsafe
        {
            get
            {
                return _loadAssembliesUnsafe;
            }
            set
            {
                if (value != _loadAssembliesUnsafe)
                {
                    _loadAssembliesUnsafe = value;
                    OnPropertyChanged("LoadAssembliesUnsafe");
                }
            }
        }

        #endregion

        #region Methods

        /// <inheritdoc/>
        public override bool Equals(object obj)
        {
            if (obj is Settings)
            {
                return this.Equals((Settings)obj);
            }

            return false;
        }

        /// <summary>
        /// Indicates whether the current object is equal to another Settings object.
        /// </summary>
        /// <param name="other">An Settings object to compare with this object.</param>
        /// <returns>true if the current object is equal to the other parameter; otherwise, false.</returns>
        public bool Equals(Settings other)
        {
            if (other == null)
                return false;

            if (other == this)
                return true;
           
            // todo: handle that better by reflection

            if (PerformanceTrace.Enabled != other.PerformanceTrace.Enabled || EnableProxyManagement != other.EnableProxyManagement ||
                EnableDynamicObjects != other.EnableDynamicObjects || EnableDynamicEventArguments != other.EnableDynamicEventArguments)
                return false;

            if ( EnableAutoDisposeEventArguments != other.EnableAutoDisposeEventArguments || EnableDynamicEventArguments != other.EnableDynamicEventArguments ||
                 ExceptionMessageBehavior != other.ExceptionMessageBehavior || ExceptionDefaultMessage != other.ExceptionDefaultMessage)
                return false;

            if (ExceptionDiagnosticsMessage != other.ExceptionDiagnosticsMessage || ThreadCulture != other.ThreadCulture ||
                EnableEvents != other.EnableEvents || MessageFilter.Enabled != other.MessageFilter.Enabled)
                return false;

            if (MessageFilter.RetryMode != other.MessageFilter.RetryMode || MessageFilter.LogMode != other.MessageFilter.LogMode ||
               EnableAutomaticQuit != other.EnableAutomaticQuit || EnableSafeMode != other.EnableSafeMode)
                return false;

            if (EnableAdHocLoading != other.EnableAdHocLoading || EnableDeepLoading != other.EnableDeepLoading ||
                 EnableMoreDebugOutput != other.EnableMoreDebugOutput || EnableEventDebugOutput != other.EnableEventDebugOutput)
                return false;

            if (CacheOptions != other.CacheOptions || EnableOperatorOverloads != other.EnableOperatorOverloads ||
                LoadAssembliesUnsafe != other.LoadAssembliesUnsafe)
                return false;

            return true;
        }

        /// <summary>
        /// Load settings from another settings instance
        /// </summary>
        /// <param name="settings">settings to load from</param>
        public void LoadFrom(Settings settings)
        {
            if (null == settings || settings == this)
                return;
            PerformanceTrace.Enabled = settings.PerformanceTrace.Enabled;
            EnableProxyManagement = settings.EnableProxyManagement;
            EnableKnownReferenceInspection = settings.EnableKnownReferenceInspection;

            EnableAutoDisposeEventArguments = settings.EnableAutoDisposeEventArguments;
            ExceptionMessageBehavior = settings.ExceptionMessageBehavior;
            ExceptionDefaultMessage = settings.ExceptionDefaultMessage;

            ExceptionDiagnosticsMessage = settings.ExceptionDiagnosticsMessage;
            ThreadCulture = settings.ThreadCulture;
            EnableEvents = settings.EnableEvents;
            MessageFilter.Enabled = settings.MessageFilter.Enabled;

            MessageFilter.RetryMode = settings.MessageFilter.RetryMode;
            MessageFilter.LogMode = settings.MessageFilter.LogMode;
            EnableAutomaticQuit = settings.EnableAutomaticQuit;
            EnableSafeMode = settings.EnableSafeMode;

            EnableAdHocLoading = settings.EnableAdHocLoading;
            EnableDeepLoading = settings.EnableDeepLoading;
            EnableMoreDebugOutput = settings.EnableMoreDebugOutput;
            EnableEventDebugOutput = settings.EnableEventDebugOutput;

            CacheOptions = settings.CacheOptions;
            EnableOperatorOverloads = settings.EnableOperatorOverloads;
            LoadAssembliesUnsafe = settings.LoadAssembliesUnsafe;
        }

        #endregion

        #region INotifyPropertyChanged

        /// <summary>
        /// Occurs when a property value changes.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion
    }
}
