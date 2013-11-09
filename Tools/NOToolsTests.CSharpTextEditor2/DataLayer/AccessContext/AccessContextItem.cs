using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    /// <summary>
    /// Access Proxy for a RootItem instance
    /// </summary>
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class AccessContextItem : INotifyPropertyChanged
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">parent list</param>
        /// <param name="dataSource">origin item</param>
        /// <param name="state">current state of the item</param>
        public AccessContextItem(AccessContextList parent, RootItem dataSource, AccessContextItemState state)
        {
            Parent = parent;
            DataSource = dataSource;
            ItemState = state;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">parent list</param>
        /// <param name="state">current state of the item</param>
        public AccessContextItem(AccessContextList parent, AccessContextItemState state)
        {
            Parent = parent;
            ItemState = state;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Parent List
        /// </summary>
        internal AccessContextList Parent { get; private set; }

        /// <summary>
        /// Returns the current state of the item instance
        /// </summary>
        internal AccessContextItemState ItemState { get; set; }

        /// <summary>
        /// Origin Item. Null if item is local new created
        /// </summary>
        internal RootItem DataSource { get; set; }

        /// <summary>
        /// Returns info the item contains local changes
        /// </summary>
        internal bool ContainsLocalChanges
        {
            get 
            {
                return LocalChangedProperties.Count > 0;
            }
        }

        /// <summary>
        /// Contains all local changed properties
        /// </summary>
        internal Dictionary<string, object> LocalChangedProperties = new Dictionary<string, object>();
        
        #endregion

        #region INotifyPropertyChanged
        
        /// <summary>
        /// Occurs when a property value has changed
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        protected internal void RaisePropertyChanged(string propertyName)
        {
            if (null != PropertyChanged)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
        
        /// <summary>
        /// Raise the ListChange event of the parent table
        /// </summary>
        protected internal void RaiseListChanged()
        {
            Parent.RaiseListChanged(ListChangedType.ItemChanged, Parent.IndexOf(this));
        }

        #endregion

        #region Methods

        /// <summary>
        /// Set a property value
        /// </summary>
        /// <param name="propertyName">name of the property</param>
        /// <param name="value">new value of the property</param>
        public void SetValue(string propertyName, object value)
        {
            object oldValue = GetValue(propertyName);
            LocalChangedProperties[propertyName] = value;
            Parent.MarkItemAsLocalChanged(this, propertyName, oldValue, value);
            RaisePropertyChanged(propertyName);
            RaiseListChanged();
        }

        /// <summary>
        /// Get a property value
        /// </summary>
        /// <param name="propertyName">name of the property</param>
        /// <returns>value of the property</returns>
        public object GetValue(string propertyName)
        {
            if (LocalChangedProperties.ContainsKey(propertyName))
                return LocalChangedProperties[propertyName];
            else
            {
                if (null != DataSource)
                    return DataSource.GetValue(propertyName);
                else
                    return null;
            }
        }

        #endregion
    }
}
