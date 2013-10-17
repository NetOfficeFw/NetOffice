using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.InMemoryCompiler
{
    /// <summary>
    /// Class definition with custom code
    /// </summary>
    public class DynamicCustomClass
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="code">custom Code</param>
        internal DynamicCustomClass(string code)
        {
            Code = code;
        }

        /// <summary>
        /// Custom code
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// Returns a string that represents the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            if (!String.IsNullOrWhiteSpace(Code))
                return Code;
            else
                return base.ToString();
        }
    }
}
