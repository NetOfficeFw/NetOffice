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
    public static class Settings
    {
        #region Imports

        [DllImport("ole32.dll", ExactSpelling = true)]
        private static extern int CoRegisterMessageFilter(IntPtr newFilter, ref IntPtr oldMsgFilter);

        #endregion

        #region Constants

        // Default Thread Culture to Excel Application
        private const string _DefaultCulture = "en-US";

        #endregion

        #region Ctor
        
        /// <summary>
        /// Static Type Constructor 
        /// </summary>
        static Settings()
        {
            _messageFilter = new RetryMessageFilter();
        }

        #endregion

        #region Fields

        private static CultureInfo  _cultureInfo;
        private static bool         _eventsEnabled = true;
        private static RetryMessageFilter  _messageFilter;
        private static bool         _enableAutomaticQuit;
        private static bool         _enableAdHocLoading = true;
        private static bool         _enableDeepLoading = true;
        private static bool         _enableDebugOutput = false;
        private static bool         _enableEventDebugOutput;
        private static bool         _enableSafeMode;
        private static bool         _enableThreadSafe = true;
        private static CacheOptions _cacheOptions = CacheOptions.KeepExistingCacheAlive;
        private static bool         _enableOperatorOverlads = true;
        private static string       _exceptionMessage = "See inner exception(s) for details.";
        private static ExceptionMessageHandling _copyInnerExceptionMessage;

        #endregion
         
        #region Properties

        /// <summary>
        /// NetOffice wrap all thrown exceptions from Office applications in a COMException. This option can be used to set the exception message
        /// </summary>
        public static ExceptionMessageHandling UseExceptionMessage
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
        /// NetOffice wrap all thrown exceptions from Office applications in a COMException. This is the default message for the top level exception 
        /// </summary>
        public static string ExceptionMessage
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
        /// Used Thread Culture given in the Invoke Calls. en-US by default
        /// </summary>
        public static CultureInfo ThreadCulture
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
                    throw (new Exception(string.Format("GetCultureInfo {0} failed.", _DefaultCulture), throwedException));
                }
                finally
                {
                    if (null == _cultureInfo)
                        throw (new Exception(string.Format("GetCultureInfo {0} failed.", _DefaultCulture)));
                }

                return _cultureInfo;
            }
            set
            {
                _cultureInfo = value;
            }
        }
        
        /// <summary>
        /// Get or set the Event support. true by default
        /// </summary>
        public static bool EnableEvents
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
        public static RetryMessageFilter MessageFilter
        {
            get
            {
                return _messageFilter;
            }
        }

        /// <summary>
        /// Get or set the Quit method for an application object was automaticly called while Dispose. false by default
        /// </summary>
        public static bool EnableAutomaticQuit
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
        public static bool EnableSafeMode
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
        /// Get or set the core api performs all operations thread safe. false by default
        /// </summary>
        public static bool EnableThreadSafe
        {
            get
            {
                return _enableThreadSafe;
            }
            set
            {
                _enableThreadSafe = value;
            }
        }

        /// <summary>
        /// Get or set Factory.Initialize() try to load non loaded dependend assemblies to fetch type informations. true by default
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static bool EnableAdHocLoading
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
        public static bool EnableDeepLoading
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
        public static bool EnableDebugOutput
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
        public static bool EnableEventDebugOutput
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
        /// Get or set Factory.Initialize() try to load non loaded dependend assemblies to fetch type informations. KeepExistingCacheAlive by default
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static CacheOptions CacheOptions
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
        public static bool EnableOperatorOverlads
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
