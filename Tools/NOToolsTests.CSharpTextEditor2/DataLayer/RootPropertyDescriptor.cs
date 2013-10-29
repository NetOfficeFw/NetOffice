using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    public class RootPropertyDescriptor : DataPropertyDescriptor
    {
        public RootPropertyDescriptor(XElement schemaNode) : base(schemaNode.Attribute("name").Value)
        {
        }

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

        public override Type ComponentType
        {
            get { return typeof(DataItem); }
        }
      
        public override bool IsReadOnly
        {
            get { return true; }
        }

        public override Type PropertyType
        {
            get { return typeof(string); }
        }

        public override bool CanResetValue(object component)
        {
            return false;
        }

        public override void ResetValue(object component)
        {
            throw new NotImplementedException();
        }
       
        public override bool ShouldSerializeValue(object component)
        {
            return false;
        }
    }
}
