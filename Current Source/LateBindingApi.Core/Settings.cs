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
        private static bool        _messageFilterEnabled;
        private static IntPtr      _messageFilter;
        private static bool        _enableAutomaticQuit;
        private static bool        _enableAdHocLoading = true;
        private static bool        _enableDebugOutput = true;

        #endregion

        #region Properties

        /// <summary>
        /// Used Thread Culture given in the Invoke Calls, default is en-US
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
        /// Get or set the Event support, default is true 
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
        /// Get or set the Message Filter enabled.
        /// The MessageFilter suspress any exceptional dialog messages, specialy the "Application ist waiting for another OLE Task" dialog
        ///default is true
        /// </summary>
        public static bool EnableMessageFilter
        {
            get
            {
                return _messageFilterEnabled;
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
                _messageFilterEnabled = value;
            }
        }

        /// <summary>
        /// Get or set the Quit() method for an application object was automaticly called while Dispose()
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
        /// Get or set Factory.Initialize() try to load non loaded dependend assemblies to fetch type informations
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
        /// Get or set additonal debug output is enabled for trouble shooting 
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
