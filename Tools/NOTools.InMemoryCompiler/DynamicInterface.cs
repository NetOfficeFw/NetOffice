using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.InMemoryCompiler
{
    /// <summary>
    /// DynamicClass interface implement
    /// </summary>
    public class DynamicInterface
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="name">Name of the interface</param>
        internal DynamicInterface(string name)
        {
            Name = name;
        }

        /// <summary>
        /// Name of the interface
        /// </summary>
        public string Name { get; private set; }
    }
}
