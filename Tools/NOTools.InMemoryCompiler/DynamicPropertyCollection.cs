using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.InMemoryCompiler
{
    /// <summary>
    /// DynamicProperty Collection
    /// </summary>
    public class DynamicPropertyCollection : List<DynamicProperty>
    {
        /// <summary>
        /// Add a property to the collection
        /// </summary>
        /// <param name="type">type of the property</param>
        /// <param name="name">name of the property</param>
        public void Add(string type, string name)
        {
            Add(new DynamicProperty(type, name));
        }

        /// <summary>
        /// returns a property instance from the collection
        /// </summary>
        /// <param name="name">name of the property</param>
        /// <returns>property instance</returns>
        public DynamicProperty this[string name]
        {
            get
            {
                foreach (var item in this)
                {
                    if (item.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase))
                        return item;
                }
                throw new ArgumentOutOfRangeException(name);
            }
        }
    }
}
