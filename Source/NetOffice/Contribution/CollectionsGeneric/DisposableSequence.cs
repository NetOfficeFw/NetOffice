using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Collections;
using NetOffice.CollectionsGeneric;

namespace NetOffice.Contribution.CollectionsGeneric
{
    /// <summary>
    /// IDisposableEnumeration'1 Default Implementation
    /// </summary>
    public class DisposableGenericList<T> : IDisposableSequence<T> where T : IDisposable
    {
        #region Fields

        private T[] _items;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="items">items to manage</param>
        public DisposableGenericList(T[] items)
        {
            _items = null != items ? items : new T[0];
        }

        #endregion

        #region IDisposable

        /// <summary>
        /// Items count of the enumeration
        /// </summary>
        public int Count
        {
            get
            {
                return _items.Length;
            }
        }

        /// <summary>
        /// Returns an index based item
        /// </summary>
        /// <param name="index">target index</param>
        /// <returns>item instance from index</returns>
        public T this[int index]
        {
            get
            {
                return _items[index];
            }
        }

        /// <summary>
        /// Returns an enumerator
        /// </summary>
        /// <returns>enumerator</returns>
        public IEnumerator<T> GetEnumerator()
        {
            foreach (T item in _items)
                yield return item;
        }

        /// <summary>
        /// Returns an enumerator
        /// </summary>
        /// <returns>enumerator</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            foreach (T item in _items)
                yield return item;
        }

        /// <summary>
        /// Dispose the instance
        /// </summary>
        /// <param name="keepAliveItem">dont dispose this item</param>
        public void Dispose(T keepAliveItem)
        {
            foreach (IDisposable item in this)
            {
                if (item.Equals(keepAliveItem))
                    continue;

                ICOMObjectDisposable disposeItem = item as ICOMObjectDisposable;
                if (null != disposeItem)
                {
                    try
                    {
                        if (false == disposeItem.IsDisposed && false == disposeItem.IsDisposed)
                            disposeItem.Dispose();
                    }
                    catch
                    {
                        ;
                    }
                }
                else
                    item.Dispose();
            }
            Clear();
        }

        /// <summary>
        /// Dispose the instance
        /// </summary>
        public void Dispose()
        {
            foreach (object item in this)
            {
                ICOMObjectDisposable disposeItem = item as ICOMObjectDisposable;
                if (null != disposeItem)
                {
                    try
                    {
                        if (false == disposeItem.IsDisposed && false == disposeItem.IsDisposed)
                            disposeItem.Dispose();
                    }
                    catch
                    {
                        ;
                    }
                }

                if (item is MarshalByRefObject)
                {
                    try
                    {
                        Marshal.ReleaseComObject(item);
                    }
                    catch
                    {
                        ;
                    }

                }
            }
            Clear();
        }

        #endregion

        #region Methods

        private void Clear()
        {
            _items = new T[0];
        }

        #endregion
    }

    /// <summary>
    /// IDisposableEnumeration Default Implementation
    /// </summary>
    public class DisposableObjectList : IDisposableSequence
    {
        #region Fields

        private object[] _items;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="items">items to manage</param>
        public DisposableObjectList(object[] items)
        {
            _items = null != items ? items : new object[0];
        }

        #endregion

        #region IDisposable

        /// <summary>
        /// Items count of the enumeration
        /// </summary>
        public int Count
        {
            get
            {
                return _items.Length;
            }
        }

        /// <summary>
        /// Returns an index based item
        /// </summary>
        /// <param name="index">target index</param>
        /// <returns>item instance from index</returns>
        public object this[int index]
        {
            get
            {
                return _items[index];
            }
        }

        /// <summary>
        /// Returns an enumerator
        /// </summary>
        /// <returns>enumerator</returns>
        public IEnumerator<object> GetEnumerator()
        {
            foreach (object item in _items)
                yield return item;
        }

        /// <summary>
        /// Returns an enumerator
        /// </summary>
        /// <returns>enumerator</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            foreach (object item in _items)
                yield return item;
        }

        /// <summary>
        /// Dispose the instance
        /// </summary>
        /// <param name="keepAliveItem">dont dispose or release this item</param>
        public void Dispose(object keepAliveItem)
        {
            foreach (object item in this)
            {
                if (item.Equals(keepAliveItem))
                    continue;

                ICOMObjectDisposable disposeItem = item as ICOMObjectDisposable;
                if (null != disposeItem)
                {
                    try
                    {
                        if (false == disposeItem.IsDisposed && false == disposeItem.IsDisposed)
                            disposeItem.Dispose();
                    }
                    catch
                    {
                        ;
                    }
                }

                if (item is MarshalByRefObject)
                {
                    try
                    {
                        Marshal.ReleaseComObject(item);
                    }
                    catch
                    {
                        ;
                    }
                }
            }
            Clear();
        }

        /// <summary>
        /// Dispose the instance
        /// </summary>
        public void Dispose()
        {
            foreach (object item in this)
            {
                ICOMObjectDisposable disposeItem = item as ICOMObjectDisposable;
                if (null != disposeItem)
                {
                    try
                    {
                        if (false == disposeItem.IsDisposed && false == disposeItem.IsDisposed)
                            disposeItem.Dispose();
                    }
                    catch
                    {
                        ;
                    }
                }

                if (item is MarshalByRefObject)
                {
                    try
                    {
                        Marshal.ReleaseComObject(item);
                    }
                    catch
                    {
                        ;
                    }

                }
            }
            Clear();
        }

        #endregion

        #region Methods

        private void Clear()
        {
            _items = new object[0];
        }

        #endregion
    }

    /// <summary>
    /// IDisposableCOMObjectSequence'1  Default Implementation
    /// </summary>
    /// <typeparam name="T">ICOMObject as any</typeparam>
    public class DisposableComObjectList<T> : IDisposableCOMObjectSequence<T> where T : ICOMObject
    {
        private T[] _items;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="items">items to manage</param>
        public DisposableComObjectList(T[] items)
        {
            _items = items;
        }


        /// <summary>
        /// Returns an index based item
        /// </summary>
        /// <param name="index">target index</param>
        /// <returns>item instance from index</returns>
        public T this[int index]
        {
            get
            {
                return _items[index];
            }
        }

        /// <summary>
        /// Items count of the enumeration
        /// </summary>
        public int Count
        {
            get
            {
                return _items.Length;
            }
        }

        /// <summary>
        /// Returns an enumerator
        /// </summary>
        /// <returns>enumerator</returns>
        public IEnumerator<T> GetEnumerator()
        {
            foreach (T item in _items)
                yield return item;
        }

        /// <summary>
        /// Returns an enumerator
        /// </summary>
        /// <returns>enumerator</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        /// <summary>
        /// Dispose the instance
        /// </summary>
        public void Dispose()
        {
            foreach (object item in this)
            {
                ICOMObjectDisposable disposeItem = item as ICOMObjectDisposable;
                if (null != disposeItem)
                {
                    if (false == disposeItem.IsDisposed && false == disposeItem.IsDisposed)
                        disposeItem.Dispose();
                }
            }
            _items = new T[0];
        }

        /// <summary>
        /// Dispose the instance
        /// </summary>
        /// <param name="keepAliveItem">dont dispose this item</param>
        public void Dispose(T keepAliveItem)
        {
            foreach (T item in this)
            {
                ICOMObjectDisposable disposeItem = item as ICOMObjectDisposable;
                if (null != disposeItem && object.ReferenceEquals(item, keepAliveItem))
                {
                    if (false == disposeItem.IsDisposed && false == disposeItem.IsDisposed)
                        disposeItem.Dispose();
                }
            }
            _items = new T[0];
        }
    }
}