using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// COMAddin want initialize Factory Core while addin startup, otherwise the Factory Core want load at first-use.
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Class, AllowMultiple = true)]
    public class ForceInitializeAttribute : System.Attribute
    {
        /// <summary>
        /// Enable Settings.EnableDebugOutput before initialize
        /// </summary>
        public readonly bool EnableDebugOutput;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public ForceInitializeAttribute()
        {

        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="enableDebugOutput">enable Settings.EnableDebugOutput before initialize</param>
        public ForceInitializeAttribute(bool enableDebugOutput)
        {
            EnableDebugOutput = enableDebugOutput;
        }
    }
}
