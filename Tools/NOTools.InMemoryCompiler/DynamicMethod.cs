using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.InMemoryCompiler
{
    /// <summary>
    /// DynamicClass Method
    /// </summary>
    public class DynamicMethod
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="name">Name of the method</param>
        /// <param name="methodCode">Source code of the method</param>
        /// <param name="returnValue">Return value of the method</param>
        internal DynamicMethod(string name, string methodCode = "", string returnValue = "")
        {
            Name = name;
            ReturnValue = returnValue;            
            MethodCode = methodCode;
            Parameters = new List<string>();
        }

        /// <summary>
        /// Name of the method
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Method arguments
        /// </summary>
        public List<String> Parameters { get; private set; }

        /// <summary>
        /// Method return value
        /// </summary>
        public string ReturnValue { get; set; }

        /// <summary>
        /// Method sourcecode
        /// </summary>
        public string MethodCode{get; set;}
    }
}
