using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Globalization;
using System.Reflection;

namespace LateBindingApi.Core
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

        #region Fields

        private static CultureInfo _cultureInfo;
        private static bool        _eventsEnabled = true;
        private static bool        _enableMessageFilter;
        private static IntPtr      _messageFilter;
        private static bool        _enableAutomaticQuit;
        private static bool        _enableAdHocLoading = true;
        private static bool        _enableDebugOutput = true;
        private static bool        _enableSafeMode;
        private static bool        _enableThreadSafe = true;

        #endregion

        #region Properties

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
        /// Get or set the Message Filter is enabled. false by default
        /// </summary>
        public static bool EnableMessageFilter
        {
            get
            {
                return _enableMessageFilter;
            }
            set
            {
                if ((value == true) && (IntPtr.Zero == _messageFilter))
                {
                    CoRegisterMessageFilter((IntPtr)0, ref _messageFilter);
                }
                else if ((value == false) && (IntPtr.Zero != _messageFilter))
                {
                    IntPtr filter = IntPtr.Zero;
                    CoRegisterMessageFilter(_messageFilter, ref filter);
                }
                _enableMessageFilter = value;
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
      
        #endregion
    }
}
