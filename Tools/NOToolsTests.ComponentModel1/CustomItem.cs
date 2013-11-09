using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOToolsTests.ComponentModel1
{
    public class CustomItem : INotifyPropertyChanged
    {
        public CustomItem()
        {

        }

        public CustomItem(string name)
        {
            Name = name;
        }

        public string Name
        {
            get
            {
                return _name;
            }
            set
            {
                _name = value;
            }
        }
        private string _name;

        public event PropertyChangedEventHandler PropertyChanged;

        private void RaisePropertyChanged(string propertyName)
        {
            if (null != PropertyChanged)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        public override string ToString()
        {
            if (!String.IsNullOrWhiteSpace(Name))
                return Name;
            else
                return "<Empty>";
        }
    }
}
