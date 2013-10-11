using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.InMemoryCompiler
{
    /// <summary>
    /// DynamicClass Collection
    /// </summary>
    public class DynamicClassCollection : List<DynamicClass>
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">Parent assembly definition</param>
        internal DynamicClassCollection(DynamicAssembly parent)
        {
            Parent = parent;
        }

        /// <summary>
        /// Parent assembly definition
        /// </summary>
        internal DynamicAssembly Parent { get; set; }

        /// <summary>
        /// Add a new item to the collection
        /// </summary>
        /// <param name="className">Name of the class</param>
        /// <param name="usings">Usings in top of the class</param>
        /// <returns>New DynamicClass instance</returns>
        public DynamicClass AddNew(string className, string[] usings = null)
        {
            DynamicClass newClass = new DynamicClass(Parent);
            Add(newClass);
            newClass.Name = className;

            if(null != usings)
                foreach (string item in usings)
                    newClass.Usings.Add(item);
            
            return newClass;
        }
    }

}
