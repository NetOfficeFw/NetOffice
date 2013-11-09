using System;
using System.Xml.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    /// <summary>
    /// Base class for specialized data property descriptor classes
    /// </summary>
    public abstract class DataPropertyDescriptor : PropertyDescriptor
    {
        #region Ctor
        
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="name">name of the property</param>
        public DataPropertyDescriptor(string name) : base(name, null)
        {
        }

        #endregion

    }
}
