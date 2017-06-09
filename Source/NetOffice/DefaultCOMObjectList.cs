using System.Collections;
using System.Collections.Generic;
using NetOffice.Interfaces;

namespace NetOffice
{
    internal class DefaultCOMObjectList : ICOMObjectList
    {
        private List<ICOMObject> _list = new List<ICOMObject>();

        public ICOMObject this[int index] { get => _list[index]; set => _list[index] = value; }

        public int Count => _list.Count;

        public bool IsReadOnly => false;

        public void Add(ICOMObject item)
        {
            _list.Add(item);
        }

        public void Clear()
        {
            _list.Clear();
        }

        public bool Contains(ICOMObject item)
        {
            return _list.Contains(item);
        }

        public void CopyTo(ICOMObject[] array, int arrayIndex)
        {
            _list.CopyTo(array, arrayIndex);
        }

        public IEnumerator<ICOMObject> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        public int IndexOf(ICOMObject item)
        {
            return _list.IndexOf(item);
        }

        public void Insert(int index, ICOMObject item)
        {
            _list.Insert(index, item);
        }

        public bool Remove(ICOMObject item)
        {
            return _list.Remove(item);
        }

        public void RemoveAt(int index)
        {
            _list.RemoveAt(index);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable)_list).GetEnumerator();
        }
    }
}
