using System;
using NetOffice.Exceptions;

namespace NetOffice
{
    /// <summary>
    /// Supports cloning, which creates a new instance of a ICOMObject instance with the same parent and underlying com proxy
    /// as an existing ICOMObject instance. Instances shares the underlying com proxy with a reference counter based lifetime system.
    /// </summary>
    /// <typeparam name="T">ICOMObject implementation</typeparam>
    /// <exception cref="CloneException">An unexpected error occured. See inner exception(s) for details.</exception>
    public interface ICloneable<T>  where T : class, ICOMObject
    {
        /// <summary>
        /// Creates a new ICOMObject that is a copy of the current instance
        /// </summary>
        /// <returns> A new ICOMObject that is a copy of this instance</returns>
        T Clone();
    }
}