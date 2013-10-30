using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.ComponentModel.InternalArray
{
    internal sealed class FunctorComparer<T> : IComparer<T>
    {
        private Comparison<T> comparison;
        private Comparer<T> c = Comparer<T>.Default;

        public FunctorComparer(Comparison<T> comparison)
        {
            this.comparison = comparison;
        }

        public int Compare(T x, T y)
        {
            return this.comparison(x, y);
        }
    }
}
