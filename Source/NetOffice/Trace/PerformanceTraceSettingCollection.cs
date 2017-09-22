using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// Collect a sequence of performance trace settings
    /// </summary>
    public class PerformanceTraceSettingCollection : List<PerformanceTraceSetting>
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        internal PerformanceTraceSettingCollection()
        {
            WildCard = new PerformanceTraceSetting("*", "*");
        }

        /// <summary>
        /// General wild card setting to trace everything
        /// </summary>
        internal PerformanceTraceSetting WildCard { get; private set; }

        /// <summary>
        /// Returns a performance trace setting by its entity name.
        /// Creates automaticaly a new performance trace setting if entity name not exists.
        /// </summary>
        /// <param name="entityName">target entity name</param>
        /// <returns>existing or new created performance trace settings</returns>
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

        /// <summary>
        /// Returns a performance trace setting by its entity name.
        /// Creates automaticaly a new performance trace setting if entity name/method name not exists.
        /// </summary>
        /// <param name="entityName">target entity name</param>
        /// <param name="methodName">target method name</param>
        /// <returns>existing or new created performance trace settings</returns>
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

        /// <summary>
        /// Get matched existing performance trace setting
        /// </summary>
        /// <param name="entityName">target entity name</param>
        /// <param name="methodName">target method</param>
        /// <returns>performance trace settings sequence</returns>
        internal IEnumerable<PerformanceTraceSetting> GetTargetEnabledSettings(string entityName, string methodName)
        {
            List<PerformanceTraceSetting> list = null;

            if (WildCard.Enabled)
            {
                list = new List<PerformanceTraceSetting>();
                list.Add(WildCard);
            }

            foreach (var item in this)
            {
                if (item.Enabled && item.EntityName == entityName && (item.MethodName == methodName || item.MethodName == "*"))
                {
                    if (null == list)
                        list = new List<PerformanceTraceSetting>();
                    list.Add(item);
                }
            }

            if (null == list)
                return new PerformanceTraceSetting[0];
            else
                return list;
        }

        /// <summary>
        /// Start time measure if entityName/methodName match
        /// </summary>
        /// <param name="entityName">target entity name</param>
        /// <param name="methodName">target method name</param>
        /// <param name="callType">invoke call kind</param>
        /// <returns>true if started, otherwise false</returns>
        internal bool TryStartMeasureTime(string entityName, string methodName, PerformanceTrace.CallType callType)
        {
            bool result = false;
            DateTime now = DateTime.Now;

            if (WildCard.Enabled)
            {
                WildCard.LastCallTime = now;
                WildCard.LastCallType = callType;
                result = true;
            }

            foreach (var item in this)
            {
                if (item.Enabled && item.EntityName == entityName && (item.MethodName == methodName || item.MethodName == "*"))
                {
                    item.LastCallTime = now;
                    item.LastCallType = callType;
                    result = true;
                }
            }

            return result;
        }
    }
}
