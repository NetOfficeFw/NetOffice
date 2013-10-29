using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    public abstract class DataPropertyDescriptor : PropertyDescriptor
    {
        public DataPropertyDescriptor(string name) : base(name, null)
        { 
        }
    }
}
