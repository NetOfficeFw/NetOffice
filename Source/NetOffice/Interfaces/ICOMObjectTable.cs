using System;
using System.Collections.Generic;
using NetOffice.Exceptions;
using System.Runtime.InteropServices;

namespace NetOffice
{
    /// <summary>
    /// Represents an <see cref="ICOMObject"/> parent/child relationship in NetOffice proxy management.
    /// </summary>
    [ComVisible(false)]
    public interface ICOMObjectTable
    {
        /// <summary>
        /// The parent object which this instance comes from or null.
        /// </summary>
        ICOMObject ParentObject { get; }

        /// <summary>
        /// Associated children objects of this instance.
        /// </summary>
        IEnumerable<ICOMObject> ChildObjects { get; }

        /// <summary>
        /// Add a new child object to this instance.
        /// </summary>
        /// <param name="childObject">new child instance</param>
        /// <exception cref="COMChildRelationException">Unexpected error</exception>
        void AddChildObject(ICOMObject childObject);

        /// <summary>
        /// Remove a child object from this instance.
        /// </summary>
        /// <param name="childObject">child instance</param>
        /// <returns>true if childObject has been removed, otherwise false</returns>
        /// <exception cref="COMChildRelationException">Unexpected error</exception>
        bool RemoveChildObject(ICOMObject childObject);
    }
}