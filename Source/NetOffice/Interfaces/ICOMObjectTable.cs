using System;
using System.Collections.Generic;
using NetOffice.Exceptions;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace NetOffice
{
    /// <summary>
    /// Represents an ICOMObject Parent/Child relation in NetOffice proxy management
    /// </summary>
    [ComVisible(false)]
    public interface ICOMObjectTable
    {
        /// <summary>
        /// The parent object where the instance come from or null(Nothing in Visual Basic)
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice")]
        ICOMObject ParentObject { get; }

        /// <summary>
        /// Associated childs from the instance
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice")]
        IEnumerable<ICOMObject> ChildObjects { get; }

        /// <summary>
        /// Add a new child to the instance
        /// </summary>
        /// <param name="childObject">new child instance</param>
        /// <exception cref="COMChildRelationException">Unexpected error</exception>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice")]
        void AddChildObject(ICOMObject childObject);

        /// <summary>
        /// Remove a child from the instance
        /// </summary>
        /// <param name="childObject">child instance</param>
        /// <returns>true if childObject has been removed, otherwise false</returns>
        /// <exception cref="COMChildRelationException">Unexpected error</exception>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice")]
        bool RemoveChildObject(ICOMObject childObject);

        /// <summary>
        /// Removes an instance from its current position in com proxy management and make him a root object
        /// </summary>
        /// <typeparam name="T">cast instance into result type</typeparam>
        /// <returns>instance result as a root proxy</returns>
        /// <exception cref="CreateInstanceException">Unexpected error</exception>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice")]
        T TakeObject<T>() where T : class, ICOMObject;
    }
}