using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Gives information about supported events
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Interface)]
    public class ComEventContractAttribute : System.Attribute
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="sinks">COM event interfaces</param>
        public ComEventContractAttribute(params Type[] sinks)
        {
            Sinks = sinks;
        }

        /// <summary>
        /// COM event interfaces
        /// </summary>
        public readonly Type[] Sinks;
    }
}
