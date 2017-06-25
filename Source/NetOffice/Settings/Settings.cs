using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Globalization;
using System.Reflection;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Threading;

namespace NetOffice
{
    /// <summary>
    /// Core Settings
    /// </summary>
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Settings
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
        private bool _enableDebugOutput = false;
        private bool _enableEventDebugOutput = false;
        private bool _enableSafeMode = false;
        private bool _enableDynamicObjects = true;

        private bool _enableProxyManagement = true;
        private CacheOptions _cacheOptions = CacheOptions.KeepExistingCacheAlive;
        private bool _enableOperatorOverlads = true;
        private string _exceptionMessage = "See inner exception(s) for details.";
        private ExceptionMessageHandling _copyInnerExceptionMessage = ExceptionMessageHandling.CopyInnerExceptionMessageToTopLevelException;
        private bool _loadAssembliesUnsafe = true;
        private PerformanceTrace _performanceTrace;
        private static Settings _default;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public Settings()
        {
            _messageFilter = new RetryMessageFilter();
            _performanceTrace = new PerformanceTrace();
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
                if (null == _default)
                    _default = new Settings();
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
                _enableProxyManagement = value;
            }
        }

        /// <summary>
        /// Wrap proxy into COMDynamicObject if proxy has no wrapper class in current app domain. true by default
        /// </summary>
        [Category("Settings"), Description("Convert unknown proxies in dynamic objects incl. proxy management."), DefaultValue(true)]
        public bool EnableDynamicObjects
        {
            get
            {
                return _enableDynamicObjects;
            }
            set
            {
                _enableDynamicObjects = value;
            }
        }

        /// <summary>
        /// NetOffice wrap all thrown exceptions from Office applications in a COMException.
        /// This option can be used to set the top level exception message or copy the innerst message to top.
        /// </summary>
        [Category("Settings"), Description("Copy inner exception message to outer top exception."), DefaultValue(typeof(ExceptionMessageHandling), "CopyInnerExceptionMessageToTopLevelException")]
        public ExceptionMessageHandling UseExceptionMessage
        {
            get
            {
                return _copyInnerExceptionMessage;
            }
            set
            {
                _copyInnerExceptionMessage = value;
            }
        }

        /// <summary>
        /// NetOffice wrap all thrown exceptions from Office applications in a COMException.
        /// This is the default message for the top level exception 
        /// </summary>
        [Category("Settings"), Description("Default exception message text."), DefaultValue("See inner exception(s) for details.")]
        public string ExceptionMessage
        {
            get
            {
                return _exceptionMessage;
            }
            set
            {
                _exceptionMessage = value;
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
                _cultureInfo = value;
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
                _eventsEnabled = value;
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
        /// Get or set the Quit method for an application object was automaticly called while Dispose. false by default
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
                _enableAutomaticQuit = value;
            }
        }

        /// <summary>
        /// Get or set the core api checks at runtime the target method or property is supported in current version. if it doesnt a EntityNotSupportedException is thrown. false by default
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
                _enableSafeMode = value;
            }
        }

        /// <summary>
        /// Get or set Core.Initialize() try to load non loaded dependend assemblies to fetch type informations. true by default
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
                _enableAdHocLoading = value;
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
                _enableDeepLoading = value;
            }
        }

        /// <summary>
        /// Get or set a flag to signalize write extended debug output
        /// </summary>
        [Category("Settings"), Description("Debug messages want be shown in console."), DefaultValue(false)]
        public bool EnableDebugOutput
        {
            get
            {
                return _enableDebugOutput;
            }
            set
            {
                _enableDebugOutput = value;
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
                _enableEventDebugOutput = value;
            }
        }

        /// <summary>
        /// Get or set Core.Initialize() try to load non loaded dependend assemblies to fetch type informations. KeepExistingCacheAlive by default
        /// </summary>
        [Category("Settings"), Description("Re-use or skip existing informations while initialize."), DefaultValue(typeof(CacheOptions), "KeepExistingCacheAlive")]
        public CacheOptions CacheOptions
        {
            get
            {
                return _cacheOptions;
            }
            set
            {
                _cacheOptions = value;
            }
        }

        /// <summary>
        /// Get or set NetOffice spend custom overloads for the "==" and "!=" operators for semanticly comparsion. true by default
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [Category("Settings"), Description("Redirect equal operations like '==' or '!=' for proxy wrapping objects to the com server to determine 2 instances are equal."), DefaultValue(true)]
        public bool EnableOperatorOverlads
        {
            get
            {
                return _enableOperatorOverlads;
            }
            set
            {
                _enableOperatorOverlads = value;
            }
        }

        /// <summary>
        /// Get or set NetOffice try load dependent assemblies unsafe(System.Reflection.Assembly.UnsafeLoadFrom). true by default
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [Category("Settings"), Description("Load assemblies unsafe and bypass some security checks."), DefaultValue(true)]
        public bool LoadAssembliesUnsafe
        {
            get
            {
                return _loadAssembliesUnsafe;
            }
            set
            {
                _loadAssembliesUnsafe = value;
            }
        }

        #endregion
    }
}
