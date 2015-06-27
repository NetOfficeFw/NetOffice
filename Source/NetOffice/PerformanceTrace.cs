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
        public bool Enabled { get; set; }

        internal string EntityName { get; private set; }

        internal string MethodName { get; private set; }

        internal DateTime LastCallTime { get; set; }

        internal PerformanceTrace.CallType LastCallType { get; set; }

        #endregion
    }

    internal class PerformanceTraceSettingCollection : List<PerformanceTraceSetting>
    {
        internal PerformanceTraceSettingCollection()
        {
            WildCard = new PerformanceTraceSetting("*", "*");
        }

        internal PerformanceTraceSetting WildCard { get; private set; }

        internal PerformanceTraceSetting this[string entityName]
        {
            get
            {
                foreach (var item in this)
                {
                    if (item.EntityName == entityName && item.MethodName == "*")
                        return item;
                }
                PerformanceTraceSetting newItem = new PerformanceTraceSetting(entityName, "*");
                Add(newItem);
                return newItem;
            }
        }

        internal PerformanceTraceSetting this[string entityName, string methodName]
        {
            get
            {
                foreach (var item in this)
                {
                    if (item.EntityName == entityName && item.MethodName == methodName)
                        return item;
                }
                PerformanceTraceSetting newItem = new PerformanceTraceSetting(entityName, methodName);
                Add(newItem);
                return newItem;
            }
        }

        internal PerformanceTraceSetting TryGetEnabledSetting(string entityName, string methodName)
        {
            if (WildCard.Enabled)
            {
                return WildCard;
            }
            else
            {
                foreach (var item in this)
                {
                    if (item.Enabled && item.EntityName == entityName && (item.MethodName == methodName || item.MethodName == "*"))
                    {
                        return item;
                    }
                }
            }

            return null;
        }

        internal bool TryStartMeasureTime(string entityName, string methodName, PerformanceTrace.CallType callType)
        {
            if (WildCard.Enabled)
            {
                WildCard.LastCallTime = DateTime.Now;
                WildCard.LastCallType = callType;
                return true;
            }
            else
            {
                foreach (var item in this)
                {
                    if (item.Enabled && item.EntityName == entityName && (item.MethodName == methodName || item.MethodName == "*") )
                    {
                        item.LastCallTime = DateTime.Now;
                        item.LastCallType = callType;
                        return true;
                    }
                }
            }
            return false;
        }
    }

    /// <summary>
    /// Call Level Performance Tracer
    /// </summary>
    public class PerformanceTrace
    {
        #region Nested

        /// <summary>
        /// Specify the kind of call
        /// </summary>
        public enum CallType
        { 
            /// <summary>
            /// Method without return value
            /// </summary>
            Method = 1,

            /// <summary>
            /// Method with return value
            /// </summary>
            Function = 2,
            
            /// <summary>
            /// Property Get
            /// </summary>
            PropertyGet = 3,

            /// <summary>
            /// Property Set
            /// </summary>
            PropertySet = 4        
        }

        /// <summary>
        /// Alert event arguments
        /// </summary>
        public class PerformanceAlertEventArgs : EventArgs
        {
            internal PerformanceAlertEventArgs(string componentName, string entityName, string methodName, double timeElapsedMS, long ticks, CallType callType, string[] arguments)
            {
                ComponentName = componentName;
                EntityName = entityName;
                MethodName = methodName;
                TimeElapsedMS = timeElapsedMS;
                Ticks = ticks;
                CallType = callType;
                Arguments = arguments;
            }

            /// <summary>
            /// Name of the corresponding NetOffice component
            /// </summary>
            public string ComponentName { get; private set; }

            /// <summary>
            /// Class name of the NetOffice wrapper
            /// </summary>
            public string EntityName { get; private set; }
            
            /// <summary>
            /// Method or property name
            /// </summary>
            public string MethodName { get; private set; }

            /// <summary>
            /// The time in milliseconds totaly
            /// </summary>
            public double TimeElapsedMS { get; private set; }

            /// <summary>
            /// The ticks totaly its need
            /// </summary>
            public long Ticks { get; private set; }

            /// <summary>
            /// Calling type
            /// </summary>
            public CallType CallType { get; private set; }

            /// <summary>
            /// Given arguments as any
            /// </summary>
            public string[] Arguments { get; private set; }
        }

        /// <summary>
        /// PerformanceTrace alert event handler
        /// </summary>
        /// <param name="sender">sender instance</param>
        /// <param name="e">alert information arguments</param>
        public delegate void PerformanceAlertEventHandler(PerformanceTrace sender, PerformanceAlertEventArgs args);

        #endregion

        #region Fields

        private object _lock;
        private Dictionary<string, PerformanceTraceSettingCollection> _repository;

        #endregion

        #region Ctor

        internal PerformanceTrace()
        {
            _lock = new object();
            _repository = new Dictionary<string, PerformanceTraceSettingCollection>();
        }

        #endregion

        #region Events

        /// <summary>
        /// Occurs if a method or property need more time as specified
        /// </summary>
        public event PerformanceAlertEventHandler Alert;

        private void RaiseAlert(string componentName, string entityName, string methodName, double timeElapsedMS, long ticks, CallType callType, string[] arguments)
        {
            if (null != Alert)
                Alert(this, new PerformanceAlertEventArgs(componentName, entityName, methodName, timeElapsedMS, ticks, callType, arguments));
        }

        #endregion

       
        #region Properties

        /// <summary>
        /// Returns performances settings instance for a NetOffice component
        /// </summary>
        /// <param name="componentName">name of the component. for example:ExcelApi</param>
        /// <returns>settings instance</returns>
        public PerformanceTraceSetting this[string componentName]
        {
            get
            {
                if (String.IsNullOrWhiteSpace(componentName))
                    throw new ArgumentNullException("componentName");

                lock (_lock)
                {
                    PerformanceTraceSettingCollection list = null;
                    if (!_repository.TryGetValue(componentName, out list))
                    {
                        list = new PerformanceTraceSettingCollection();
                        _repository.Add(componentName, list);
                    }

                    return list.WildCard;
                }
            }
        }

        /// <summary>
        /// Returns performance settings instance for a NetOffice wrapper class
        /// </summary>
        /// <param name="componentName">name of the component. for example:ExcelApi</param>
        /// <param name="entityName">name of the class. for example:Range or Application</param>
        /// <returns>settings instance</returns>
        public PerformanceTraceSetting this[string componentName, string entityName]
        {
            get
            {
                if (String.IsNullOrWhiteSpace(componentName))
                    throw new ArgumentNullException("componentName");
                if (String.IsNullOrWhiteSpace(entityName))
                    throw new ArgumentNullException("entityName");
             
                lock (_lock)
                {
                    PerformanceTraceSettingCollection list = null;
                    if (!_repository.TryGetValue(componentName, out list))
                    {
                        list = new PerformanceTraceSettingCollection();
                        _repository.Add(componentName, list);
                    }

                    return list[entityName];
                }
            }
        }

        /// <summary>
        /// Returns performance settings instance for a NetOffice wrapper class
        /// </summary>
        /// <param name="componentName">name of the component. for example:ExcelApi</param>
        /// <param name="entityName">name of the class. for example:Range or Application</param>
        /// <param name="methodName">method or property name. for example: Visible or Activate</param>
        /// <returns>settings instance</returns>
        public PerformanceTraceSetting this[string componentName, string entityName, string methodName]
        {
            get
            {
                if (String.IsNullOrWhiteSpace(componentName))
                    throw new ArgumentNullException("componentName");
                if (String.IsNullOrWhiteSpace(entityName))
                    throw new ArgumentNullException("entityName");
                if (String.IsNullOrWhiteSpace(methodName))
                    throw new ArgumentNullException("methodName");

                lock (_lock)
                {
                    PerformanceTraceSettingCollection list = null;
                    if (!_repository.TryGetValue(componentName, out list))
                    {
                        list = new PerformanceTraceSettingCollection();
                        _repository.Add(componentName, list);
                    }

                    return list[entityName, methodName];
                }
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Clear all performance trace settings
        /// </summary>
        public void Clear()
        {
            lock (_lock)
            {
                _repository.Clear();                
            }
        }

        internal bool StartMeasureTime(string componentName, string entityName, string methodName, CallType callType)
        {
            PerformanceTraceSettingCollection list = null;
            if (_repository.TryGetValue(componentName, out list))
                return list.TryStartMeasureTime(entityName, methodName, callType);
            else
                return false;
        }

        internal void StopMeasureTime(string componentName, string entityName, string methodName, params object[] arguments)
        {
            DateTime now = DateTime.Now;
            PerformanceTraceSettingCollection list = null;
            if (_repository.TryGetValue(componentName, out list))
            {
                PerformanceTraceSetting setting = list.TryGetEnabledSetting(entityName, methodName);
                if (null != setting)
                {
                    TimeSpan ts = now - setting.LastCallTime;
                    if (ts.TotalMilliseconds >= setting.IntervalMS)
                    {
                        List<string> args = new List<string>();
                        foreach (var item in arguments)
                            args.Add((null == item || item == Type.Missing) ? "<Empty>" : item.ToString());
                        RaiseAlert(componentName, entityName, methodName, ts.TotalMilliseconds, ts.Ticks, setting.LastCallType, args.ToArray());

                    }
                }
            }
        }

        #endregion
    }
}
