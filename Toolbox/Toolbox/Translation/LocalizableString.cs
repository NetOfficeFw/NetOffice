using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.Translation
{
    /// <summary>
    /// Represents a localizable single-line string
    /// </summary>
    internal class LocalizableString : NotifyPropertyChanged
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        internal LocalizableString()
        { 
                  
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="value">name</param>
        internal LocalizableString(string value)
        {
            _value = value;
        }

        /// <summary>
        /// Creates an instance of the class 
        /// </summary>
        /// <param name="value">name</param>
        /// <param name="value2">localized value</param>
        internal LocalizableString(string value, string value2)
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
            return new LocalizableString(_value, _value2);
        }
    }
}
