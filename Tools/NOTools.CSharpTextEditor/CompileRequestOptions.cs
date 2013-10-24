using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace NOTools.CSharpTextEditor
{
    /// <summary>
    /// Options for user compile request
    /// </summary>
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class CompileRequestOptions
    {
        public CompileRequestOptions()
        {
            CompileRequestKey = Key.F6;
            RunRequestKey = Key.F5;
        }

        /// <summary>
        /// Enables the request mode
        /// </summary>
        [Description("Enables the request mode"), Category("Options")]
        public bool Enabled { get; set; }
        
        /// <summary>
        /// Target F key for request
        /// </summary>
        [Description("Enables the request mode"), Category("Options")]
        public Key CompileRequestKey { get; set; }
        
        /// <summary>
        /// Target F key for request
        /// </summary>
        [Description("Enables the request mode"), Category("Options")]
        public Key RunRequestKey { get; set; }

        /// <summary>
        /// Returns a System.String that represents the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
           return "CompileRequest";
        }
    }
}
