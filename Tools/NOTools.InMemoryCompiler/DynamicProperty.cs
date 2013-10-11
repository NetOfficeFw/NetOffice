using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.InMemoryCompiler
{
    /// <summary>
    /// DynamicClass Property 
    /// </summary>
    public class DynamicProperty
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="type">Type of the property</param>
        /// <param name="name">Name of the property</param>
        internal DynamicProperty(string type, string name)
        {
            Type = type;
            Name = name;
        }

        /// <summary>
        /// Type of the property ("String" or "System.Xml.Linq.XDocument" for example)
        /// </summary>
        public string Type { get; set; }

        /// <summary>
        /// Name of the property ("MyProperty" for example)
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// System.String that represents the current instance
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return String.Format("{0} {1}", Type, Name).Trim();
        }
    }
}
