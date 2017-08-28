using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// Performance Trace Options
    /// </summary>
    public class PerformanceTraceSetting
    {
        #region Fields

        private int _intervalMS = 1000;
        private bool _enabled = false;

        #endregion

        #region Ctor

        internal PerformanceTraceSetting(string entityName, string methodName)
        {
            EntityName = entityName;
            MethodName = methodName;
        }

        internal PerformanceTraceSetting(string entityName, string methodName, int intervalMS)
        {
            EntityName = entityName;
            MethodName = methodName;
            IntervalMS = intervalMS;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Alert limit in milliseconds. Default:1000
        /// If a calling method or property need more(or equal) time as specified here, the alert event is fired
        /// </summary>
        public int IntervalMS
        {
            get
            {
                return _intervalMS;
            }
            set
            {
                if (value < 0)
                    value = 0;
                _intervalMS = value;
            }
        }

        /// <summary>
        /// Enable or disable trace alert
        /// </summary>
        public bool Enabled
        {
            get
            {
                return _enabled;
            }
            set
            {
                if (_enabled != value)
                    _enabled = value;
            }
        }

        internal string EntityName { get; private set; }

        internal string MethodName { get; private set; }

        internal DateTime LastCallTime { get; set; }

        internal PerformanceTrace.CallType LastCallType { get; set; }

        #endregion
    }
}
