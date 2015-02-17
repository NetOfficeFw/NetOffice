using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.Translation
{
    /// <summary>
    /// Represents a localizable multiline string
    /// </summary>
    internal class LocalizableWideString : NotifyPropertyChanged
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        internal LocalizableWideString()
        { 
                  
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="value">name</param>
        internal LocalizableWideString(string value)
        {
            _value = value;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="value">origin value</param>
        /// <param name="value2">name</param>
        internal LocalizableWideString(string value, string value2)
        {
            _value = value;
            _value2 = value2;
        }

        /// <summary>
        /// Create a deep copy of the instance
        /// </summary>
        /// <returns>clone instance</returns>
        public override NotifyPropertyChanged Clone()
        {
            return new LocalizableWideString(_value, _value2);
        }
    }
}
