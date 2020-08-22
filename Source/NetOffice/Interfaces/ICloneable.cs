using System;
using NetOffice.Exceptions;

namespace NetOffice
{
    /// <summary>
    /// Supports cloning, which creates a new instance of a <see cref="ICOMObject"/> instance with the same parent and underlying com proxy
    /// as an existing <see cref="ICOMObject"/> instance. Instances shares the underlying com proxy with a reference counter based lifetime system.
    /// </summary>
    /// <typeparam name="T"><see cref="ICOMObject"/> implementation</typeparam>
    /// <exception cref="CloneException">An unexpected error occured. See inner exception(s) for details.</exception>
    public interface ICloneable<out T>  where T : class, ICOMObject
    {
        /// <summary>
        /// Creates a new <see cref="ICOMObject"/> that is a copy of the current instance
        /// </summary>
        /// <returns> A new <see cref="ICOMObject"/> that is a copy of this instance</returns>
        T Clone();
    }
}