using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// Represents a Parent/Child relation 
    /// </summary>
    public interface ICOMObjectTable
    {
        /// <summary>
        /// The parent object where the instance come from or null(Nothing in Visual Basic)
        /// </summary>
        ICOMObject ParentObject { get; }

        /// <summary>
        /// Associated childs from the instance
        /// </summary>
        IEnumerable<ICOMObject> ChildObjects { get; }

        /// <summary>
        /// Add a new child to the instance
        /// </summary>
        /// <param name="childObject">new child instance</param>
        void AddChildObject(ICOMObject childObject);

        /// <summary>
        /// Remove a child from the instance
        /// </summary>
        /// <param name="childObject">child instance</param>
        void RemoveChildObject(ICOMObject childObject);
    }

}
