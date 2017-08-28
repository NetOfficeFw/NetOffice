using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Gives information about connection sink implementations to support events
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Interface)]
    public class EventSinkAttribute : System.Attribute
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="sinks">sinks based on NetOffice.SinkHelper</param>
        public EventSinkAttribute(params Type[] sinks)
        {
            Sinks = sinks;
        }

        /// <summary>
        /// Sinks based on NetOffice.SinkHelper
        /// </summary>
        public readonly Type[] Sinks;
    }
}