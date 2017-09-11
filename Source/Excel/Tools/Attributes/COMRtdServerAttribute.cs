using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.ExcelApi.Tools.Attributes
{
    /// <summary>
    /// Class is marked as RTD Server in NetOffice
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class COMRtdServerAttribute : System.Attribute
    {
        /// <summary>
        /// Name of the Rtd server
        /// </summary>
        public readonly string Name;

        /// <summary>
        /// Default heartbeat of the Rtd server.
        /// Default is 1 if COMRtdServerAttribute is not set and Heartbeat is not overriden.
        /// </summary>
        public readonly int Heartbeat;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public COMRtdServerAttribute() : this(String.Empty)
        {
            
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="heartbeat">default heartbeat</param>
        public COMRtdServerAttribute(int heartbeat) : base()
        {
            Heartbeat = heartbeat;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="name">given name as any</param>
        public COMRtdServerAttribute(string name) : base()
        {
            Name = name;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="name">given name as any</param>
        /// <param name="heartbeat">default heartbeat</param>
        public COMRtdServerAttribute(string name, int heartbeat) : base()
        {
            Name = name;
            Heartbeat = heartbeat;
        }
    }
}
