using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    /// <summary>
    /// Base class for specialized classes that represents a table row.
    /// </summary>
    public abstract class DataItem : INotifyPropertyChanged
    {
        #region Fields

        /// <summary>
        /// Item Properties
        /// </summary>
        protected internal Dictionary<string, object> Properties = new Dictionary<string, object>();
        
        #endregion

        #region Methods

        /// <summary>
        /// Set a property value
        /// </summary>
        /// <param name="propertyName">name of the property</param>
        /// <param name="value">new value for the property</param>
        public virtual void SetValue(string propertyName, object value)
        {
            Properties[propertyName] = value;
            RaisePropertyChanged(propertyName);
        }

        /// <summary>
        /// Returns a property value
        /// </summary>
        /// <param name="propertyName">name of the property</param>
        /// <returns>property value (null if not exists)</returns>
        public virtual object GetValue(string propertyName)
        {
            object result = null;
            Properties.TryGetValue(propertyName, out result);
            return result;
        }

        #endregion

        #region INotifyPropertyChanged

        /// <summary>
        /// Occures when a property value has changed
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        protected internal void RaisePropertyChanged(string propertyName)
        {
            if (null != PropertyChanged)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion
    }
}
