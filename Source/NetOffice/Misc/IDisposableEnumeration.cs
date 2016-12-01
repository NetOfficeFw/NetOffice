using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;
using System.Text;
using System.Collections;
using System.ComponentModel;

namespace NetOffice.Misc
{
    /// <summary>
    /// Represents an IEnumerable:T with the service of you can dispose all items with one call.
    /// NOTE: IDisposable.Dispose() want ignore any ObjectDisposedException and want clear the collection
    /// at the end if no unexpected error occurs.
    /// After Dispose the instance want call ReleaseComObject for each item if its a COM proxy.
    /// </summary>
    public interface IDisposableEnumeration<T>: IDisposable, IEnumerable<T> where T : IDisposable 
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
    }
    
    /// <summary>
    /// Represents an IEnumerable"object" with the service of you can dispose all items with one call.
    /// NOTE: IDisposable.Dispose() want ignore any ObjectDisposedException and want clear the collection
    /// at the end if no unexpected error occurs.
    /// After Dispose the instance want call ReleaseComObject if its a COM proxy.
    /// </summary>
    public interface IDisposableEnumeration : IDisposable, IEnumerable<object>
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
    } 
}

namespace NetOffice.Misc
{
    /// <summary>
    /// IDisposableEnumeration Default Implementation
    /// </summary>
    public class DisposableObjectList : List<object>, IDisposableEnumeration
    {
        #region IDisposable

        /// <summary>
        /// Dispose the instance
        /// </summary>
        public void Dispose()
        {
            foreach (object item in this)
            {
                IDisposable disposeItem = item as IDisposable;
                if (null != disposeItem)
                {
                    try
                    {
                        disposeItem.Dispose();
                    }
                    catch (ObjectDisposedException)
                    {
                        // not nice but i didnt find a way to check an IDisposable instance is already disposed 
                        // - SL
                        ;
                    }
                    catch
                    {
                        throw;
                    }

                    if (item is MarshalByRefObject)
                    {
                        Marshal.ReleaseComObject(item);
                    }
                }
            }
            Clear();
        }

        #endregion
    }
}