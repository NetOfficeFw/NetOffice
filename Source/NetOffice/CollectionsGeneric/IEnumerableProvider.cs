using System;
using System.Collections;
using System.Collections.Generic;
using NetOffice;
using NetOffice.Exceptions;

namespace NetOffice.CollectionsGeneric
{
    /// <summary>
    /// Provides enumerable sequence services
    /// </summary>
    /// <typeparam name="T">T as any</typeparam>
    public interface IEnumerableProvider<T> : IEnumerable<T>
    {
        /// <summary>
        /// Creates a managed enumerator
        /// </summary>
        /// <param name="parent">parent instance or null in com proxy management</param>
        /// <returns>managed enumerator</returns>
        /// <exception cref="ArgumentNullException">argument is null</exception>
        /// <exception cref="NetOfficeCOMException">unexpected error</exception>
        ICOMObject GetComObjectEnumerator(ICOMObject parent);

        /// <summary>
        /// Fetch managed enumerator
        /// </summary>
        /// <param name="parent">parent instance or null in com proxy management</param>
        /// <param name="enumerator">enumerator to fetch</param>
        /// <returns>IEnumerator instance</returns>
        /// <exception cref="ArgumentNullException">argument is null</exception>
        /// <exception cref="NetOfficeCOMException">unexpected error</exception>
        IEnumerable FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator);
    }
}