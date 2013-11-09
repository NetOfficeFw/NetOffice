using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    /// <summary>
    /// Property Descriptor for RootList.cs
    /// </summary>
    public class RootPropertyDescriptor : DataPropertyDescriptor
    {
        #region Ctor
        
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="schemaNode">schema defintion</param>
        public RootPropertyDescriptor(XElement schemaNode) : base(schemaNode.Attribute("name").Value)
        {
            SchemaNode = schemaNode;
        }

        #endregion

        #region Properties
        
        /// <summary>
        /// Schema Definition
        /// </summary>
        protected internal XElement SchemaNode { get; private set; }

        #endregion

        #region Overrides

        public override object GetValue(object component)
        {
            DataItem item = component as DataItem;
            if (null != item)
                return item.GetValue(this.Name);
            else
                return null;
        }

        public override void SetValue(object component, object value)
        {
            DataItem item = component as DataItem;
            if (null != item)
                item.SetValue(this.Name, value);
        }

        public override bool CanResetValue(object component)
        {
            return false;
        }

        public override Type ComponentType
        {
            get
            {
                return typeof(DataItem);
            }
        }

        public override bool IsReadOnly
        {
            get
            {
                return Convert.ToBoolean(SchemaNode.Attribute("readonly").Value);
            }
        }

        public override Type PropertyType
        {
            get
            {
                return Type.GetType(SchemaNode.Attribute("type").Value);
            }
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
