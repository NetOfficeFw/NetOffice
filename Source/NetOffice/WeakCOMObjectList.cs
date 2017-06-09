using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.Interfaces;

namespace NetOffice
{
    internal sealed class WeakCOMObjectList : ICOMObjectList
    {
        private List<WeakReference> _list = new List<WeakReference>();

        public ICOMObject this[int index]
        {
            get => (ICOMObject)_list[index].Target;
            set => _list[index] = new WeakReference(value);
        }

        public int Count
        {
            get
            {
                Update();
                return _list.Count;
            }
        }

        public bool IsReadOnly => false;

        public void Add(ICOMObject item)
        {
            Update();
            _list.Add(new WeakReference(item));
        }

        public void Clear()
        {
            _list.Clear();
        }

        public bool Contains(ICOMObject item)
        {
            return _list.Any(wr => wr.Target == item);
        }

        public void CopyTo(ICOMObject[] array, int arrayIndex)
        {
            Update();
            _list.Select(wr => wr.Target).ToList().CopyTo(array, arrayIndex);
        }

        public IEnumerator<ICOMObject> GetEnumerator()
        {
            Update();
            return _list.Select(wr => wr.Target).Cast<ICOMObject>().GetEnumerator();
        }

        public int IndexOf(ICOMObject item)
        {
            throw new NotImplementedException();
        }

        public void Insert(int index, ICOMObject item)
        {
            throw new NotImplementedException();
        }

        public bool Remove(ICOMObject item)
        {
            Update();
            for (var i = 0; i < _list.Count; i++)
            {
                if (_list[i].Target == item)
                {
                    _list.RemoveAt(i);
                    return true;
                }
            }

            return false;
        }

        public void RemoveAt(int index)
        {
            _list.RemoveAt(index);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            Update();
            return _list.Select(wr => wr.Target).Cast<ICOMObject>().GetEnumerator();
        }

        private void Update()
        {
            // This implementation uses logic similar to List<T>.RemoveAll, which always has O(n) time.
            //  Some other implementations seen in the wild have O(n*m) time, where m is the number of dead entries.
            //  As m approaches n (e.g., mass object extinctions), their running time approaches O(n^2).
            // See https://github.com/StephenCleary/Mvvm.Core/blob/master/src/Nito.Mvvm.Core/WeakCollection.cs.
            int writeIndex = 0;
            for (int readIndex = 0; readIndex != _list.Count; ++readIndex)
            {
                var weakReference = _list[readIndex];
                if (weakReference.Target != null)
                {
                    if (readIndex != writeIndex)
                    {
                        _list[writeIndex] = _list[readIndex];
                    }

                    ++writeIndex;
                }
            }

            _list.RemoveRange(writeIndex, _list.Count - writeIndex);
        }
    }
}
