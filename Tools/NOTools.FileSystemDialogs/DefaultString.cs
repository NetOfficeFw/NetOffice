using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class DefaultString
    {
        public bool UseDefault { get; set; }
        public string Value { get; set; }
    }
}
