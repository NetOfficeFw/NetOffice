using System;
using System.Collections;
using System.Collections.Generic;

namespace NetOffice.CollectionsGeneric
{
    /// <summary>
    /// Represents an IEnumerable:T with the service of disposing all items
    /// </summary>
    public interface IDisposableCOMObjectSequence<T> : IDisposable, IEnumerable<T> where T : ICOMObject
    {
        /// <summary>
        /// Returns an index based item
        /// </summary>
        /// <param name="index">target index</param>
        /// <returns>item instance from index</returns>
        T this[int index] { get; }

        /// <summary>
        /// Items count of the enumeration
        /// </summary>
        int Count { get; }
        
        /// <summary>
        /// Dispose the instance
        /// </summary>
        /// <param name="keepAliveItem">dont dispose this item</param>
        void Dispose(T keepAliveItem);
    }

    /// <summary>
    /// Represents an IEnumerable:T with the service of disposing all items
    /// </summary>
    public interface IDisposableSequence<T>: IDisposable, IEnumerable<T> where T : IDisposable 
    {
        /// <summary>
        /// Returns an index based item
        /// </summary>
        /// <param name="index">target index</param>
        /// <returns>item instance from index</returns>
        T this[int index] { get; }

        /// <summary>
        /// Items count of the enumeration
        /// </summary>
        int Count { get; }

        /// <summary>
        /// Dispose the instance
        /// </summary>
        /// <param name="keepAliveItem">dont dispose this item</param>
        void Dispose(T keepAliveItem);
    }

    /// <summary>
    /// Represents an IEnumerable:T with the service of disposing all items.
    /// IDisposableEnumeration want dispose items there implement IDispose and
    /// call Marshal.ReleaseComObject if item is a Com Proxy.
    /// </summary>
    public interface IDisposableSequence : IDisposable, IEnumerable<object>
    {
        /// <summary>
        /// Returns an index based item
        /// </summary>
        /// <param name="index">target index</param>
        /// <returns>item instance from index</returns>
        object this[int index] { get; }

        /// <summary>
        /// Items count of the enumeration
        /// </summary>
        int Count { get; }

        /// <summary>
        /// Dispose the instance
        /// </summary>
        /// <param name="keepAliveItem">dont dispose or release this item</param>
        void Dispose(object keepAliveItem);
    }
}