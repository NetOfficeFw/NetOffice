using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.InMemoryCompiler
{
    /// <summary>
    /// Usings Collection
    /// </summary>s
    public class DynamicUsingsCollection : List<string>
    {
        /// <summary>
        /// Creates an instance ot the class
        /// </summary>
        public DynamicUsingsCollection()
        {
            Add("System");
            Add("System.Windows.Forms");
        }

        /// <summary>
        /// Add an item to the collection
        /// </summary>
        /// <param name="item">Using item</param>
        public new void Add(string item)
        { 
            if(!(this.Contains(item)))
                base.Add(item);
        }
    }
}
