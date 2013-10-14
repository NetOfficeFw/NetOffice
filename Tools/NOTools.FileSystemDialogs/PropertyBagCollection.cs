using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    internal delegate void PropertyChangedHandler(string message);

    /// <summary>
    /// Dictionary with DefaultValue
    /// </summary>
    /// <typeparam name="TKey">key</typeparam>
    /// <typeparam name="TValue">value</typeparam>
    public class PropertyBagCollection<TValue> : Dictionary<string, TValue>
    {
        #region Ctor
        
        /// <summary>
        /// Creates am instance of the class
        /// </summary>
        /// <param name="defaultValue">Default for non existing key-value pairs</param>
        internal PropertyBagCollection(TValue defaultValue, PropertyChangedHandler propertyChangedHandler = null, KeyValuePair<string, TValue>[] initialValues = null)
        {        
            DefaultValue = defaultValue;
            PropertyChangedHandler = propertyChangedHandler;
            if (null != initialValues)
                foreach (KeyValuePair<string, TValue> item in initialValues)
                    this[item.Key] = item.Value;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Target default valaue
        /// </summary>
        internal TValue DefaultValue { get; set; }

        /// <summary>
        /// Atached PropertyChanged Event
        /// </summary>
        internal PropertyChangedHandler PropertyChangedHandler { get; set; }

        /// <summary>
        /// Get or set a value with a key
        /// </summary>
        /// <param name="key">key</param>
        /// <returns>value</returns>
        public new TValue this[string key]
        {
            get
            {
                TValue outValue = default(TValue);
                if (base.TryGetValue(key, out outValue))
                    return outValue;
                else
                    return DefaultValue;
            }
            set
            {
                base[key] = value;
                if (null != PropertyChangedHandler)
                    PropertyChangedHandler(key);
            }
        }

        #endregion
    }
}
