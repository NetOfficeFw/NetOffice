using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.InMemoryCompiler
{
    /// <summary>
    /// References Collection 
    /// </summary>
    public class DynamicReferencesCollection : List<string>
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public DynamicReferencesCollection()
        {
            //Add("mscorlib.dll");
            //Add("System.dll");
            //Add("System.Windows.Forms.dll");
        }

        /// <summary>
        /// Add a new reference
        /// </summary>
        /// <param name="item">new reference</param>
        public new void Add(string item)
        {
            if (!(this.Contains(item)))
                base.Add(item);
        }
    }
}
