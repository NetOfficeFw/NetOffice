using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    /// <summary>
    /// PropertyDescriptor for AccessContextList instances
    /// </summary>
    public class AccessContextPropertyDescriptor : PropertyDescriptor
    {
        #region Ctor
        
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="dataSourcePropertyDescriptor">origin descriptor from RootList instance</param>
        public AccessContextPropertyDescriptor(PropertyDescriptor dataSourcePropertyDescriptor) : base(dataSourcePropertyDescriptor.Name, null)
        {
            DataSourcePropertyDescriptor = dataSourcePropertyDescriptor;
        }

        #endregion

        #region Properties
        
        /// <summary>
        /// Origin Descriptor from RootList instance
        /// </summary>
        internal PropertyDescriptor DataSourcePropertyDescriptor { get; private set; }
        
        public override object GetValue(object component)
        {
            AccessContextItem item = component as AccessContextItem;
            if (null != item)
                return item.GetValue(Name);
            else
                return null;
        }

        public override void SetValue(object component, object value)
        {
            AccessContextItem item = component as AccessContextItem;
            if (null != item)
                item.SetValue(Name, value);
        }

        public override bool CanResetValue(object component)
        {
            return false;
        }

        public override Type ComponentType
        {
            get { return DataSourcePropertyDescriptor.ComponentType; }
        }

        public override bool IsReadOnly
        {
            get { return DataSourcePropertyDescriptor.IsReadOnly; }
        }

        public override Type PropertyType
        {
            get { return DataSourcePropertyDescriptor.PropertyType; }
        }

        public override void ResetValue(object component)
        {
            throw new NotImplementedException();
        }

        public override bool ShouldSerializeValue(object component)
        {
            return false;
        }

        #endregion
    }
}
