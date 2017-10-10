using System;
using System.Collections.Generic;
using NetOffice;
using NetOffice.Exceptions;
using NetOffice.CollectionsGeneric;
using System.Collections;

namespace NetOffice.Extensions
{
    /// <summary>
    /// Provides a set of static (Shared in Visual Basic) methods for querying objects
    /// that implement NetOffice.CollectionsGeneric.IEnumerableProvider`1.
    /// </summary>
    public static class EnumerableExtensions
    {   
        /// <summary>
        /// Returns the first element of a sequence
        /// </summary>
        /// <typeparam name="TSource">the type of the elements of <paramref name="source"/></typeparam>
        /// <param name="source">the <see cref="T:System.Collections.Generic.IEnumerable`1" /> to return the first element of</param>
        /// <param name="append">append items in sequence to parent instance in com proxy management</param>
        /// <returns>the first element in the specified sequence</returns>
        /// <exception cref="ArgumentNullException">source is null(Nothing in Visual Basic)</exception>
        /// <exception cref="InvalidOperationException">sequence is empty</exception>
        /// <exception cref="NetOfficeCOMException">error occured while calling remote server</exception>
        public static TSource First<TSource>(this IEnumerableProvider<TSource> source, bool append = true)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            ICOMObject enumerator = null;
            try
            {
                enumerator = source.GetComObjectEnumerator(null);
                foreach (TSource item in source.FetchVariantComObjectEnumerator(true == append ? source as ICOMObject : null, enumerator))
                {
                    if (null != enumerator && enumerator != source)
                        enumerator.Dispose();
                    return item;
                }
                throw new InvalidOperationException("Sequence is empty.");
            }
            catch
            {
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                throw;
            }
        }
        
        /// <summary>
        /// Returns the first element of a sequence
        /// </summary>
        /// <typeparam name="TSource">the type of the elements of <paramref name="source"/></typeparam>
        /// <param name="source">the <see cref="T:System.Collections.Generic.IEnumerable`1" /> to return the first element of</param>
        /// <param name="append">append items in sequence to parent instance in com proxy management</param>
        /// <param name="predicate">a function to test each element for a condition</param>
        /// <returns>the first element in the specified sequence</returns>
        /// <exception cref="ArgumentNullException">source is null(Nothing in Visual Basic)</exception>
        /// <exception cref="InvalidOperationException">sequence is empty</exception>
        /// <exception cref="NetOfficeCOMException">error occured while calling remote server</exception>
        public static TSource First<TSource>(this IEnumerableProvider<TSource> source, Func<TSource, bool> predicate, bool append = true)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            ICOMObject enumerator = null;
            try
            {
                IEnumerableProvider<TSource> sequence = source as IEnumerableProvider<TSource>;
                if (null == sequence)
                    throw new ArgumentException("Unable to cast IEnumerableProvider<TSource>");
                enumerator = sequence.GetComObjectEnumerator(null);
                foreach (TSource item in sequence.FetchVariantComObjectEnumerator(true == append ? source as ICOMObject : null, enumerator))
                {
                    if (predicate(item))
                    {
                        if (null != enumerator && enumerator != source)
                            enumerator.Dispose();
                        return item;
                    }
                    else
                    {
                        TryDispose(item);
                    }
                }
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                throw new InvalidOperationException("No element satisfies the condition in predicate.-or-The source sequence is empty.");
            }
            catch
            {
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Returns the first element of a sequence, or a default value if the sequence contains no elements
        /// </summary>
        /// <typeparam name="TSource">the type of the elements of <paramref name="source"/></typeparam>
        /// <param name="source">the <see cref="T:System.Collections.Generic.IEnumerable`1" /> to return the first element of</param>
        /// <param name="append">append items in sequence to parent instance in com proxy management</param>
        /// <returns>default(TSource) if <paramref name="source" /> is empty; otherwise, the first element in <paramref name="source" /></returns>
        /// <exception cref="ArgumentNullException">source is null(Nothing in Visual Basic)</exception>
        /// <exception cref="NetOfficeCOMException">error occured while calling remote server</exception>     
        public static TSource FirstOrDefault<TSource>(this IEnumerableProvider<TSource> source, bool append = true)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            ICOMObject enumerator = null;
            try
            {
                enumerator = source.GetComObjectEnumerator(null);
                foreach (TSource item in source.FetchVariantComObjectEnumerator(true == append ? source as ICOMObject : null, enumerator))
                {
                    if (null != enumerator && enumerator != source)
                        enumerator.Dispose();
                    return item;
                }
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                return default(TSource);
            }
            catch
            {
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Returns the first element of a sequence, or a default value if the sequence contains no elements
        /// </summary>
        /// <typeparam name="TSource">the type of the elements of <paramref name="source"/></typeparam>
        /// <param name="source">the <see cref="T:System.Collections.Generic.IEnumerable`1" /> to return the first element of</param>
        /// <param name="predicate">a function to test each element for a condition</param>
        /// <param name="append">append items in sequence to parent instance in com proxy management</param>
        /// <returns>default(TSource) if <paramref name="source" /> is empty; otherwise, the first element in <paramref name="source" /></returns>
        /// <exception cref="ArgumentNullException">source is null(Nothing in Visual Basic)</exception>
        /// <exception cref="NetOfficeCOMException">error occured while calling remote server</exception> 
        public static TSource FirstOrDefault<TSource>(this IEnumerableProvider<TSource> source, Func<TSource, bool> predicate, bool append = true)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            ICOMObject enumerator = null;
            try
            {
                enumerator = source.GetComObjectEnumerator(null);
                foreach (TSource item in source.FetchVariantComObjectEnumerator(true == append ? source as ICOMObject : null, enumerator))
                {
                    if (predicate(item))
                    {
                        if (null != enumerator && enumerator != source)
                            enumerator.Dispose();
                        return item;
                    }
                    else
                        TryDispose(item);
                }
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                return default(TSource);
            }
            catch
            {
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Returns the last element of a sequence
        /// </summary>
        /// <typeparam name="TSource">the type of the elements of <paramref name="source"/></typeparam>
        /// <param name="source">the <see cref="T:System.Collections.Generic.IEnumerable`1" /> to return the first element of</param>
        /// <param name="append">append items in sequence to parent instance in com proxy management</param>
        /// <returns>the value at the last position in the source sequence</returns>
        /// <exception cref="ArgumentNullException">source is null(Nothing in Visual Basic)</exception>
        /// <exception cref="InvalidOperationException">sequence is empty</exception>
        /// <exception cref="NetOfficeCOMException">error occured while calling remote server</exception>
        public static TSource Last<TSource>(this IEnumerableProvider<TSource> source, bool append = true)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            ICOMObject enumerator = null;
            try
            {
                enumerator = source.GetComObjectEnumerator(null);
                IEnumerator enumeratorSource = source.FetchVariantComObjectEnumerator(true == append ? source as ICOMObject : null, enumerator).GetEnumerator();
                if (enumeratorSource.MoveNext())
                {
                    TSource lastCurrent = default(TSource);
                    TSource current = default(TSource);
                    do
                    {
                        lastCurrent = current;
                        current = (TSource)enumeratorSource.Current;
                        if (null != lastCurrent)
                            TryDispose(lastCurrent);
                    }
                    while (enumeratorSource.MoveNext());
                    if (null != enumerator && enumerator != source)
                        enumerator.Dispose();
                    if (null != current)
                        return current;
                }
                throw new InvalidOperationException("Sequence is empty.");
            }
            catch
            {
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Returns the last element of a sequence
        /// </summary>
        /// <typeparam name="TSource">the type of the elements of <paramref name="source"/></typeparam>
        /// <param name="source">the <see cref="T:System.Collections.Generic.IEnumerable`1" /> to return the first element of</param>
        /// <param name="append">append items in sequence to parent instance in com proxy management</param>
        /// <param name="predicate">a function to test each element for a condition</param>
        /// <returns>the value at the last position in the source sequence</returns>
        /// <exception cref="ArgumentNullException">source is null(Nothing in Visual Basic)</exception>
        /// <exception cref="InvalidOperationException">sequence is empty</exception>
        /// <exception cref="NetOfficeCOMException">error occured while calling remote server</exception>
        public static TSource Last<TSource>(this IEnumerableProvider<TSource> source, Func<TSource, bool> predicate, bool append = true)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            ICOMObject enumerator = null;
            try
            {
                enumerator = source.GetComObjectEnumerator(null);
                IEnumerator enumeratorSource = source.FetchVariantComObjectEnumerator(true == append ? source as ICOMObject : null, enumerator).GetEnumerator();
                if (enumeratorSource.MoveNext())
                {
                    TSource lastCurrent = default(TSource);
                    TSource current = default(TSource);
                    do
                    {
                        lastCurrent = current;
                        TSource item = (TSource)enumeratorSource.Current;
                        if (predicate(item))
                            current = item;
                        if (null != lastCurrent)
                            TryDispose(lastCurrent);
                    }
                    while (enumeratorSource.MoveNext());
                    if (null != enumerator && enumerator != source)
                        enumerator.Dispose();
                    if (null != current)
                        return current;
                }
                throw new InvalidOperationException("Sequence is empty.");
            }
            catch
            {
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                throw;
            }
        }
       
        /// <summary>
        /// Returns the last element of a sequence
        /// </summary>
        /// <typeparam name="TSource">the type of the elements of <paramref name="source"/></typeparam>
        /// <param name="source">the <see cref="T:System.Collections.Generic.IEnumerable`1" /> to return the first element of</param>
        /// <param name="append">append items in sequence to parent instance in com proxy management</param>
        /// <returns>the value at the last position in the source sequence</returns>
        /// <exception cref="ArgumentNullException">source is null(Nothing in Visual Basic)</exception>
        /// <exception cref="NetOfficeCOMException">error occured while calling remote server</exception>
        public static TSource LastOrDefault<TSource>(this IEnumerableProvider<TSource> source, bool append = true)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            ICOMObject enumerator = null;
            try
            {
                enumerator = source.GetComObjectEnumerator(null);
                IEnumerator enumeratorSource = source.FetchVariantComObjectEnumerator(true == append ? source as ICOMObject : null, enumerator).GetEnumerator();
                if (enumeratorSource.MoveNext())
                {
                    TSource lastCurrent = default(TSource);
                    TSource current = default(TSource);
                    do
                    {
                        lastCurrent = current;
                        current = (TSource)enumeratorSource.Current;
                        if (null != lastCurrent)
                            TryDispose(lastCurrent);
                    }
                    while (enumeratorSource.MoveNext());
                    if (null != current)
                        return current;
                }
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                return default(TSource);
            }
            catch
            {
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Returns the last element of a sequence
        /// </summary>
        /// <typeparam name="TSource">the type of the elements of <paramref name="source"/></typeparam>
        /// <param name="source">the <see cref="T:System.Collections.Generic.IEnumerable`1" /> to return the first element of</param>
        /// <param name="append">append items in sequence to parent instance in com proxy management</param>
        /// <param name="predicate">a function to test each element for a condition</param>
        /// <returns>the value at the last position in the source sequence</returns>
        /// <exception cref="ArgumentNullException">source is null(Nothing in Visual Basic)</exception>
        /// <exception cref="NetOfficeCOMException">error occured while calling remote server</exception>
        public static TSource LastOrDefault<TSource>(this IEnumerableProvider<TSource> source, Func<TSource, bool> predicate, bool append = true)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            ICOMObject enumerator = null;
            try
            {
                enumerator = source.GetComObjectEnumerator(null);
                IEnumerator enumeratorSource = source.FetchVariantComObjectEnumerator(true == append ? source as ICOMObject : null, enumerator).GetEnumerator();
                if (enumeratorSource.MoveNext())
                {
                    TSource lastCurrent = default(TSource);
                    TSource current = default(TSource);
                    do
                    {
                        lastCurrent = current;
                        TSource item = (TSource)enumeratorSource.Current;
                        if (predicate(item))
                            current = item;
                        if (null != lastCurrent)
                            TryDispose(lastCurrent);
                    }
                    while (enumeratorSource.MoveNext());
                    if (null != current)
                        return current;
                }
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                return default(TSource);
            }
            catch
            {
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Returns the number of elements in a sequence
        /// </summary>
        /// <typeparam name="TSource">the type of the elements of <paramref name="source"/></typeparam>
        /// <param name="source">the <see cref="T:System.Collections.Generic.IEnumerable`1" /> to return the first element of</param>
        /// <param name="append">append items in sequence to parent instance in com proxy management</param>
        /// <returns>the number of elements in the input sequence.</returns>
        /// <exception cref="ArgumentNullException">source is null(Nothing in Visual Basic)</exception>
        /// <exception cref="NetOfficeCOMException">error occured while calling remote server</exception>
        public static int Count<TSource>(this IEnumerableProvider<TSource> source, bool append = true)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            int result = 0;
            ICOMObject enumerator = null;
            try
            {
                enumerator = source.GetComObjectEnumerator(null);
                foreach (TSource current in source.FetchVariantComObjectEnumerator(true == append ? source as ICOMObject : null, enumerator))
                {
                    result++;
                    TryDispose(current);
                }
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
            }
            catch
            {
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                throw;
            }
            return result;
        }

        /// <summary>
        /// Returns the number of elements in a sequence
        /// </summary>
        /// <typeparam name="TSource">the type of the elements of <paramref name="source"/></typeparam>
        /// <param name="source">the <see cref="T:System.Collections.Generic.IEnumerable`1" /> to return the first element of</param>
        /// <param name="predicate">a function to test each element for a condition</param>
        /// <param name="append">append items in sequence to parent instance in com proxy management</param>
        /// <returns>the number of elements in the input sequence.</returns>
        /// <exception cref="ArgumentNullException">source is null(Nothing in Visual Basic)</exception>
        /// <exception cref="NetOfficeCOMException">error occured while calling remote server</exception>
        public static int Count<TSource>(this IEnumerableProvider<TSource> source, Func<TSource, bool> predicate, bool append = true)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            int result = 0;
            ICOMObject enumerator = null;
            try
            {
                enumerator = source.GetComObjectEnumerator(null);
                foreach (TSource current in source.FetchVariantComObjectEnumerator(true == append ? source as ICOMObject : null, enumerator))
                {
                    if(predicate(current))
                        result++;
                    TryDispose(current);
                }
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
            }
            catch
            {
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                throw;
            }
            return result;
        }

        /// <summary>
        /// Determines whether a sequence contains any elements
        /// </summary>
        /// <typeparam name="TSource">the type of the elements of <paramref name="source"/></typeparam>
        /// <param name="source">the <see cref="T:System.Collections.Generic.IEnumerable`1" /> to return the first element of</param>
        /// <param name="append">append items in sequence to parent instance in com proxy management</param>
        /// <returns>true if the source sequence contains any elements; otherwise, false</returns>
        /// <exception cref="ArgumentNullException">source is null(Nothing in Visual Basic)</exception>
        /// <exception cref="NetOfficeCOMException">error occured while calling remote server</exception>
        public static bool Any<TSource>(this IEnumerableProvider<TSource> source, bool append = true)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            ICOMObject enumerator = null;
            try
            {
                IEnumerableProvider<TSource> sequence = source as IEnumerableProvider<TSource>;
                if (null == sequence)
                    throw new ArgumentException("Unable to cast IEnumerableProvider<TSource>");
                enumerator = sequence.GetComObjectEnumerator(null);
                foreach (TSource current in sequence.FetchVariantComObjectEnumerator(true == append ? source as ICOMObject : null, enumerator))
                {
                    TryDispose(current);
                    return true;
                }

                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                return false;
            }
            catch
            {
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Determines whether a sequence contains any elements
        /// </summary>
        /// <typeparam name="TSource">the type of the elements of <paramref name="source"/></typeparam>
        /// <param name="source">the <see cref="T:System.Collections.Generic.IEnumerable`1" /> to return the first element of</param>
        /// <param name="predicate">a function to test each element for a condition</param>
        /// <param name="append">append items in sequence to parent instance in com proxy management</param>
        /// <returns>true if the source sequence contains any elements; otherwise, false</returns>
        /// <exception cref="ArgumentNullException">source is null(Nothing in Visual Basic)</exception>
        /// <exception cref="NetOfficeCOMException">error occured while calling remote server</exception>
        public static bool Any<TSource>(this IEnumerableProvider<TSource> source, Func<TSource, bool> predicate, bool append = true)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            ICOMObject enumerator = null;
            try
            {
                enumerator = source.GetComObjectEnumerator(null);
                foreach (TSource current in source.FetchVariantComObjectEnumerator(true == append ? source as ICOMObject : null, enumerator))
                {
                    bool match = predicate(current);
                    TryDispose(current);
                    if (match)
                        return true;
                }

                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                return false;
            }
            catch
            {
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Determines whether a sequence contains a specified element by using the NetOffice Core equality comparer
        /// </summary>
        /// <typeparam name="TSource">the type of the elements of <paramref name="source"/></typeparam>
        /// <param name="source">the <see cref="T:System.Collections.Generic.IEnumerable`1" /> to return the first element of</param>
        /// <param name="value">the value to locate in the sequence</param>
        /// <param name="append">append items in sequence to parent instance in com proxy management</param>
        /// <returns>true if the source sequence contains an element that has the specified value; otherwise, false</returns>
        /// <exception cref="ArgumentNullException">source is null(Nothing in Visual Basic)</exception>
        /// <exception cref="NetOfficeCOMException">error occured while calling remote server</exception>
        public static bool Contains<TSource>(this IEnumerableProvider<TSource> source, TSource value, bool append = true)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            ICOMObject enumerator = null;
            try
            {
                enumerator = source.GetComObjectEnumerator(null);
                foreach (TSource current in source.FetchVariantComObjectEnumerator(true == append ? source as ICOMObject : null, enumerator))
                {
                    bool match = Core.EqualsOnServer(current, value);
                    TryDispose(current);
                    if (match)
                        return true;
                }

                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                return false;
            }
            catch
            {
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Filters a sequence of values based on a predicate
        /// </summary>
        /// <typeparam name="TSource">the type of the elements of <paramref name="source"/></typeparam>
        /// <param name="source">the <see cref="T:System.Collections.Generic.IEnumerable`1" /> to return the first element of</param>
        /// <param name="predicate">a function to test each element for a condition</param>
        /// <param name="append">append items in sequence to parent instance in com proxy management</param>
        /// <returns>true if the source sequence contains an element that has the specified value; otherwise, false</returns>
        /// <exception cref="ArgumentNullException">source is null(Nothing in Visual Basic)</exception>
        /// <exception cref="NetOfficeCOMException">error occured while calling remote server</exception>
        public static IEnumerable<TSource> Where<TSource>(this IEnumerableProvider<TSource> source, Func<TSource, int, bool> predicate, bool append = true)
        {
            if (null == source)
                throw new ArgumentNullException("source");
            if (null == predicate)
                throw new ArgumentNullException("predicate");
            ICOMObject enumerator = null;
            List<TSource> list = new List<TSource>();
            try
            {
                enumerator = source.GetComObjectEnumerator(null);
                int num = -1;
                foreach (TSource item in source.FetchVariantComObjectEnumerator(true == append ? source as ICOMObject : null, enumerator))
                {
                    int num2 = num;
                    num = checked(num2 + 1);
                    if (predicate(item, num))
                        list.Add(item);
                    else
                        TryDispose(item);
                }
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                return list;
            }
            catch
            {               
                for (int i = list.Count - 1; i >= 0; i++)
                    TryDispose(list[i]);
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Filters a sequence of values based on a predicate
        /// </summary>
        /// <typeparam name="TSource">the type of the elements of <paramref name="source"/></typeparam>
        /// <param name="source">the <see cref="T:System.Collections.Generic.IEnumerable`1" /> to return the first element of</param>
        /// <param name="predicate">a function to test each element for a condition</param>
        /// <param name="append">append items in sequence to parent instance in com proxy management</param>
        /// <returns>true if the source sequence contains an element that has the specified value; otherwise, false</returns>
        /// <exception cref="ArgumentNullException">source is null(Nothing in Visual Basic)</exception>
        /// <exception cref="NetOfficeCOMException">error occured while calling remote server</exception>
        public static IEnumerable<TSource> Where<TSource>(this IEnumerableProvider<TSource> source, Func<TSource, bool> predicate, bool append = true)
        {
            if (null == source)
                throw new ArgumentNullException("source");
            if (null == predicate)
                throw new ArgumentNullException("predicate");
            ICOMObject enumerator = null;
            List<TSource> list = new List<TSource>();
            try
            {
                enumerator = source.GetComObjectEnumerator(null);
                foreach (TSource item in source.FetchVariantComObjectEnumerator(true == append ? source as ICOMObject : null, enumerator))
                {
                    if (predicate(item))
                        list.Add(item);
                    else
                        TryDispose(item);
                }
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                return list;
            }
            catch
            {
                for (int i = list.Count-1; i >=0; i++)
                    TryDispose(list[i]);
                if (null != enumerator && enumerator != source)
                    enumerator.Dispose();
                throw;
            }
        }

        private static void TryDispose(object item)
        {
            ICOMObject comObject = item as ICOMObject;
            if (null != comObject)
                comObject.Dispose();
        }
    }
}