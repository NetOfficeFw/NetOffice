using System;
using System.Collections.Generic;
using NetOffice.Exceptions;

namespace NetOffice
{
    /// <summary>
    /// Represents an ICOMObject Parent/Child relation in NetOffice proxy management
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
        /// <exception cref="COMChildRelationException">Unexpected error</exception>
        void AddChildObject(ICOMObject childObject);

        /// <summary>
        /// Remove a child from the instance
        /// </summary>
        /// <param name="childObject">child instance</param>
        /// <returns>true if childObject has been removed, otherwise false</returns>
        /// <exception cref="COMChildRelationException">Unexpected error</exception>
        bool RemoveChildObject(ICOMObject childObject);
    }
}