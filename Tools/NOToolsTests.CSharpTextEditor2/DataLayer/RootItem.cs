using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    /// <summary>
    /// 
    /// </summary>
    public class RootItem : DataItem
    {
        Dictionary<string, object> _rowValues = new Dictionary<string, object>();

        public RootItem()
        {
        }

        public RootItem(XElement dataNode)
        {
            foreach (var item in dataNode.Attributes())
                _rowValues.Add(item.Name.LocalName, item.Value);
        }

        public override void SetValue(string propertyName, object value)
        {
            if (_rowValues.ContainsKey(propertyName))
                _rowValues[propertyName] = value;
        }

        public override object GetValue(string propertyName)
        {
            return _rowValues[propertyName];     
        }
    }
}
