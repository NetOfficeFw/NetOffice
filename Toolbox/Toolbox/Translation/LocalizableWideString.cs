using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.Translation
{
    internal class LocalizableWideString : NotifyPropertyChanged
    {
        internal LocalizableWideString()
        { 
                  
        }

        internal LocalizableWideString(string value)
        {
            _value = value;
        }

        internal LocalizableWideString(string value, string value2)
        {
            _value = value;
            _value2 = value2;
        }

        public override NotifyPropertyChanged Clone()
        {
            return new LocalizableWideString(_value, _value2);
        }
    }
}
