using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    public abstract class DataItem : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public abstract void SetValue(string propertyName, object value);

        public abstract object GetValue(string propertyName);
    }
}
