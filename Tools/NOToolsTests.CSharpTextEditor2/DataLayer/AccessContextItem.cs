using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class AccessContextItem : DataItem , INotifyPropertyChanged
    {
        public override void SetValue(string propertyName, object value)
        {
            throw new NotImplementedException();
        }

        public override object GetValue(string propertyName)
        {
            throw new NotImplementedException();
        }
    }
}
