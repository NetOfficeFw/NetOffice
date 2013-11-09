using System;
using System.IO;
using System.Reflection;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Text;

using CM = NOTools.ComponentModel;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    /// <summary>
    /// Base class for specialized classes that represents a table or a list property.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public abstract class DataList<T> : CM.BindingList<T>, ITypedList where T : DataItem
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="listName">name of the list</param>
        public DataList(string listName)
        {
            Name = listName;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Name of the list
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Indicates the list is already loaded from database
        /// </summary>
        public virtual bool IsLoaded { get; protected internal set; }

        #endregion

        #region Methods

        /// <summary>
        /// Load data from database or any other source. 
        /// The method has to set the IsLoaded property to true when its loaded successfully.
        /// </summary>
        public abstract void LoadFromDatabase();

        /// <summary>
        /// Returns property descriptors the list instance
        /// </summary>
        /// <param name="listAccessors">target properties. unused argument!</param>
        /// <returns>property descriptors</returns>
        public abstract PropertyDescriptorCollection GetItemProperties(PropertyDescriptor[] listAccessors);

        /// <summary>
        /// Returns the name of the list
        /// </summary>
        /// <param name="listAccessors">target properties. unused argument!</param>
        /// <returns>System.String</returns>
        public virtual string GetListName(PropertyDescriptor[] listAccessors)
        {
            return Name;
        }

        /// <summary>
        /// Returns a System.String instance that represents the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            return "DataList " + Name;
        }

        #endregion
    }
}
