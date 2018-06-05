using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.CoreExtensions
{
    /// <summary>
    /// Provides type convert extension
    /// </summary>
    public static class CoreTypeExtensions
    {
        /// <summary>
        /// Analyze an object and create wrapper arround if necessary
        /// </summary>
        /// <param name="factory">core to extend</param>
        /// <param name="value">value as any</param>
        /// <param name="allowDynamicObject">allow to create a COMDynamicObject instance if its failed to resolve the wrapper type</param>
        /// <returns>value or wrapped value</returns>
        public static object WrapObject(this Core factory, object value, bool allowDynamicObject)
        {
            if ((null != value) && (value is MarshalByRefObject))
            {
                ICOMObject newObject = factory.CreateObjectFromComProxy(null, value, allowDynamicObject);
                return newObject;
            }
            else
            {
                return value;
            }
        }
    }
}
