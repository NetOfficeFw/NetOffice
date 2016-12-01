using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Misc
{
    /// <summary>
    /// BindingList with generic sort
    /// </summary>
    /// <typeparam name="T">generic argument</typeparam>
    public class SortableBindingList<T> : BindingList<T>, ITypedList where T : class
    {
        #region Fields

        private bool _isSorted;
        private ListSortDirection _sortDirection = ListSortDirection.Ascending;
        private PropertyDescriptor _sortProperty;

        #endregion
        
        #region Ctor

        /// <summary>
        /// Initializes a new instance of the <see cref="SortableBindingList{T}"/> class.
        /// </summary>
        public SortableBindingList()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SortableBindingList{T}"/> class.
        /// </summary>
        /// <param name="list">An <see cref="T:System.Collections.Generic.IList`1" /> of items to be contained in the <see cref="T:System.ComponentModel.BindingList`1" />.</param>
        public SortableBindingList(IList<T> list) : base(list)
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets a value indicating whether the list supports sorting.
        /// </summary>
        protected override bool SupportsSortingCore
        {
            get { return true; }
        }

        /// <summary>
        /// Gets a value indicating whether the list is sorted.
        /// </summary>
        protected override bool IsSortedCore
        {
            get { return _isSorted; }
        }

        /// <summary>
        /// Gets the direction the list is sorted.
        /// </summary>
        protected override ListSortDirection SortDirectionCore
        {
            get { return _sortDirection; }
        }

        /// <summary>
        /// Gets the property descriptor that is used for sorting the list if sorting is implemented in a derived class; otherwise, returns null
        /// </summary>
        protected override PropertyDescriptor SortPropertyCore
        {
            get { return _sortProperty; }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Removes any sort applied with ApplySortCore if sorting is implemented
        /// </summary>
        protected override void RemoveSortCore()
        {
            _sortDirection = ListSortDirection.Ascending;
            _sortProperty = null;
            _isSorted = false;
        }

        /// <summary>
        /// Sorts the items if overridden in a derived class
        /// </summary>
        /// <param name="prop"></param>
        /// <param name="direction"></param>
        protected override void ApplySortCore(PropertyDescriptor prop, ListSortDirection direction)
        {
            _sortProperty = prop;
            _sortDirection = direction;

            List<T> list = Items as List<T>;
            if (list == null)
                return;

            list.Sort(Compare);

            _isSorted = true;
            OnListChanged(new ListChangedEventArgs(ListChangedType.Reset, -1));
        }

        private int Compare(T lhs, T rhs)
        {
            int result = OnComparison(lhs, rhs);

            if (_sortDirection == ListSortDirection.Descending)
                result = -result;
            return result;
        }

        private int OnComparison(T lhs, T rhs)
        {
            object lhsValue = lhs == null ? null : _sortProperty.GetValue(lhs);
            object rhsValue = rhs == null ? null : _sortProperty.GetValue(rhs);
            if (lhsValue == null)
            {
                return (rhsValue == null) ? 0 : -1;
            }
            if (rhsValue == null)
            {
                return 1;
            }
            if (lhsValue is IComparable)
            {
                return ((IComparable)lhsValue).CompareTo(rhsValue);
            }
            if (lhsValue.Equals(rhsValue))
            {
                return 0;
            }
          
            return lhsValue.ToString().CompareTo(rhsValue.ToString());
        }


        #endregion

        #region ITypedList

        /// <summary>
        /// Returns always String.Empty
        /// </summary>
        /// <param name="listAccessors">An array of System.ComponentModel.PropertyDescriptor objects, for which the list name is returned. This can be null.</param>
        /// <returns>The name of the list</returns>
        public virtual string GetListName(PropertyDescriptor[] listAccessors)
        {
            return String.Empty;
        }

        /// <summary>
        /// Returns the System.ComponentModel.PropertyDescriptorCollection that represents
        /// the properties on each item used to bind data.
        /// </summary>
        /// <param name="listAccessors">An array of System.ComponentModel.PropertyDescriptor objects, for which the list name is returned. This can be null.</param>
        /// <returns>PropertyDescriptorCollection that represents the properties</returns>
        public virtual PropertyDescriptorCollection GetItemProperties(PropertyDescriptor[] listAccessors)
        {
            return TypeDescriptor.GetProperties(typeof(T));
        }

        #endregion
    }
}
