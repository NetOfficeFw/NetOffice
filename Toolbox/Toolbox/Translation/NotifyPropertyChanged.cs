using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.Translation
{
    public class NotifyPropertyChanged : INotifyPropertyChanged
    {
        protected string _value;
        protected string _value2;

        public string Value
        {
            get
            {
                return _value;
            }
            set
            {
                _value = value;
                RaisePropertyChanged("Value");
            }
        }
       
        public string Value2
        {
            get
            {
                return _value2;
            }
            set
            {
                _value2 = value;
                RaisePropertyChanged("Value2");
            }
        }

        public virtual NotifyPropertyChanged Clone()
        {
            throw new NotSupportedException();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected internal void RaisePropertyChanged(string propertyName)
        {
            if (null != PropertyChanged)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        public override string ToString()
        {
            return String.Format("Value:{0}   Value2:{1}", _value, _value2);                
        }
    }
}
