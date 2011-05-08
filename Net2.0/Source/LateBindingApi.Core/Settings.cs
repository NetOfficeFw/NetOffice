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
        private static bool        _eventsEnabled;
        private static bool        _messageFilterEnabled;
        private static IntPtr      _messageFilter;

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
        /// Get or set the Event support, default is false 
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

        #endregion
    }
}
