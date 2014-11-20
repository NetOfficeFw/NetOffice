using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.Translation
{
    internal class LocalizableRTFString : NotifyPropertyChanged
    {
        internal LocalizableRTFString()
        { 
                  
        }

        internal LocalizableRTFString(string value)
        {
            _value = value;
        }

        internal LocalizableRTFString(string value, string value2)
        {
            _value = value;
            _value2 = value2;
        }

        public override NotifyPropertyChanged Clone()
        {
            return new LocalizableRTFString(_value, _value2);
        }
    }
}
