using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.InMemoryCompiler
{
    /// <summary>
    /// DynamicMethod Collection
    /// </summary>
    public class DynamicMethodCollection : List<DynamicMethod>
    {       
        /// <summary>
        ///  Add a new item to the collection
        /// </summary>
        /// <param name="name">Name of the method</param>
        /// <param name="methodCode">Source Code of the method</param>
        /// <param name="returnValue">Return value of the method</param>
        /// <returns>New DynamicMethod instance</returns>
        public DynamicMethod AddNew(string name, string methodCode = "", string returnValue = "")
        {
            DynamicMethod newMethode = new DynamicMethod(name, methodCode, returnValue);
            Add(newMethode);
            return newMethode;
        }

        /// <summary>
        /// Returns an item of the collection
        /// </summary>
        /// <param name="methodName">Name of the method</param>
        /// <returns>DynamicMethod instance</returns>
        public DynamicMethod this[string methodName]
        {
            get
            {
                foreach (var item in this)
                    if (item.Name.Equals(methodName, StringComparison.InvariantCultureIgnoreCase))
                        return item;
                throw new ArgumentOutOfRangeException(methodName);
            }
        }
    }
}
