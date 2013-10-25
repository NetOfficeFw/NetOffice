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
    public class Settings
    {
        #region Constants

        // Default Thread Culture to Excel Application
        private const string _DefaultCulture = "en-US";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public Settings()
        {
            _messageFilter = new RetryMessageFilter();
        }

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
        private bool _enableProxyManagement = true;
        private bool _enableProxyCountChannel = false;
        private string _proxyCountChannelName = "DefaultProxyCountChannel";
        private CacheOptions _cacheOptions = CacheOptions.KeepExistingCacheAlive;
        private bool _enableOperatorOverlads = true;
        private string _exceptionMessage = "See inner exception(s) for details.";
        private ExceptionMessageHandling _copyInnerExceptionMessage;

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
        private static Settings _default;

        /// <summary>
        /// Enable the NetOffice COM proxy management. true by default
        /// </summary>
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
        /// NetOffice wrap all thrown exceptions from Office applications in a COMException.
        /// This option can be used to set the top level exception message or copy the innerst message to top.
        /// </summary>
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
                    throw new ArgumentNullException("ThreadCulture must have a value");
                _cultureInfo = value;
            }
        }

        /// <summary>
        /// Get or set the Event support. true by default
        /// </summary>
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
        /// Get or set Factory.Initialize() try to load non loaded dependend assemblies to fetch type informations. true by default
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
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
        /// Get or set the Initialize method perform a deep level analyzing(may cause security issues)
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
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
        /// Get or set NetOffice logs essential system steps in the DebugConsole(if enabled). true by default
        /// </summary>
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
        /// Get or set NetOffice logs the current proxy count from a core factory to a ConsoleMonitor channel
        /// </summary>
        public bool EnableProxyCountChannel
        {
            get
            {
                return _enableProxyCountChannel;
            }
            set
            {
                _enableProxyCountChannel = value;
            }
        }

        /// <summary>
        /// A ConsoleMonitor channel name for the option EnableProxyCountChannel 
        /// </summary>
        public string ProxyCountChannelName
        {
            get
            {
                return _proxyCountChannelName;
            }
            set
            {
                _proxyCountChannelName = value;
            }
        }

        /// <summary>
        /// Get or set Factory.Initialize() try to load non loaded dependend assemblies to fetch type informations. KeepExistingCacheAlive by default
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
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

        #endregion
    }
}