using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.InMemoryCompiler
{
    /// <summary>
    /// DynamicCustomClass Collection
    /// </summary>
    public class DynamicCustomClassCollection : List<DynamicCustomClass>
    {
        /// <summary>
        /// Add a new item to the collection
        /// </summary>
        /// <param name="code">custom code for new DynamicCustomClass instance</param>
        /// <returns>new item instance</returns>
        public DynamicCustomClass AddNew(string code)
        {
            DynamicCustomClass newItem = new DynamicCustomClass(code);
            base.Add(newItem);
            return newItem;
        }
    }
}
