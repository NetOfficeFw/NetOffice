using System;
using System.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.Collections;

namespace NetOffice
{
    /// <summary>
    /// Call Level Performance Tracer
    /// </summary>
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class PerformanceTrace : IEnumerable<KeyValuePair<string, PerformanceTraceSettingCollection>>
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

            /// <summary>
            /// Returns a System.String that represents the instance
            /// </summary>
            /// <returns>System.String</returns>
            public override string ToString()
            {
                return String.Format("{0} {1} {2}, {3} Milliseconds", ComponentName, EntityName, MethodName, TimeElapsedMS);
            }
        }

        /// <summary>
        /// PerformanceTrace alert event handler
        /// </summary>
        /// <param name="sender">sender instance</param>
        /// <param name="args">alert information arguments</param>
        public delegate void PerformanceAlertEventHandler(PerformanceTrace sender, PerformanceAlertEventArgs args);

        #endregion

        #region Fields

        private object _lock;
        private Dictionary<string, PerformanceTraceSettingCollection> _repository;
        private bool _enabled;

        #endregion

        #region Ctor

        internal PerformanceTrace()
        {
            _lock = new object();
            _repository = new Dictionary<string, PerformanceTraceSettingCollection>();
        }

        internal PerformanceTrace(Action<string> onPropertyChanged)
        {
            _lock = new object();
            _repository = new Dictionary<string, PerformanceTraceSettingCollection>();
            OnPropertyChanged = onPropertyChanged;
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
        /// Enable or disable the performance trace system
        /// </summary>
        [Description("Enable or disable the performance trace system"), DefaultValue(false), Category("PerformanceTrace")]
        public bool Enabled
        {
            get
            {
                return _enabled;
            }
            set
            {
                if (value != _enabled)
                { 
                    _enabled = value;
                    OnPropertyChanged?.Invoke("PerformanceTrace.Enabled");
                }
            }
        }

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
                        OnPropertyChanged?.Invoke("PerformanceTrace.Item");
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
                        OnPropertyChanged?.Invoke("PerformanceTrace.Item");
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
                        OnPropertyChanged?.Invoke("PerformanceTrace.Item");
                    }

                    return list[entityName, methodName];
                }
            }
        }
        
        /// <summary>
        /// Occurs when a property value changes
        /// </summary>
        private Action<string> OnPropertyChanged { get; set; }

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
            if (!Enabled)
                return false;

            PerformanceTraceSettingCollection list = null;
            if (_repository.TryGetValue(componentName, out list))
            {
               
                bool result = list.TryStartMeasureTime(entityName, methodName, callType);
                return result;
            }
            else
                return false;
        }

        internal void StopMeasureTime(string componentName, string entityName, string methodName, params object[] arguments)
        {
            DateTime now = DateTime.Now;
            PerformanceTraceSettingCollection list = null;
            if (_repository.TryGetValue(componentName, out list))
            {
                IEnumerable<PerformanceTraceSetting> settings = list.GetTargetEnabledSettings(entityName, methodName);
                foreach (var item in settings)
                {
                    TimeSpan ts = now - item.LastCallTime;
                    if (ts.TotalMilliseconds >= item.IntervalMS)
                    {
                        List<string> args = new List<string>();
                        foreach (var arg in arguments)
                            args.Add((null == arg || arg == Type.Missing) ? "<Empty>" : arg.ToString());
                        RaiseAlert(componentName, entityName, methodName, ts.TotalMilliseconds, ts.Ticks, item.LastCallType, args.ToArray());
                        return;
                    }
                }
            }
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// Sequence of all traces
        /// </summary>
        /// <returns>sequence</returns>
        public IEnumerator<KeyValuePair<string, PerformanceTraceSettingCollection>> GetEnumerator()
        {
            return _repository.GetEnumerator();
        }

        /// <summary>
        /// Sequence of all traces
        /// </summary>
        /// <returns>sequence</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _repository.GetEnumerator();
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Returns a System.String that represents the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            return String.Format("{0} Trace(s)", _repository.Count);
        }

        #endregion
    }
}
