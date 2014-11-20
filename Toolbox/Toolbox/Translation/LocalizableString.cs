using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.Translation
{
    internal class LocalizableString : NotifyPropertyChanged
    {
        internal LocalizableString()
        { 
                  
        }

        internal LocalizableString(string value)
        {
            _value = value;
        }

        internal LocalizableString(string value, string value2)
        {
            _value = value;
            _value2 = value2;
        }

        public override NotifyPropertyChanged Clone()
        {
            return new LocalizableString(_value, _value2);
        }
    }
}
