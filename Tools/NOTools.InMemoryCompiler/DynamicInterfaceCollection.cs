using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.InMemoryCompiler
{
    /// <summary>
    /// DynamicInterface Colllection
    /// </summary>
    public class DynamicInterfaceCollection : List<DynamicInterface>
    {
        /// <summary>
        /// Add an item to the collection
        /// </summary>
        /// <param name="className">Name of the interfac</param>
        /// <returns>new item instance</returns>
        internal DynamicInterface AddNew(string interfaceName)
        {
            DynamicInterface newInterface = new DynamicInterface(interfaceName);
            Add(newInterface);
            return newInterface;
        }
    }

}
