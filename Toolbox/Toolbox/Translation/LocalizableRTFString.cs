using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.Translation
{
    /// <summary>
    /// Represents a localizable rich text string
    /// </summary>
    internal class LocalizableRTFString : NotifyPropertyChanged
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        internal LocalizableRTFString()
        { 
                  
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="value">name</param>
        internal LocalizableRTFString(string value)
        {
            _value = value;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="value">name</param>
        /// <param name="value2">value</param>
        internal LocalizableRTFString(string value, string value2)
        {
            _value = value;
            _value2 = value2;
        }

        /// <summary>
        /// Create deep copy of the instance
        /// </summary>
        /// <returns>clone</returns>
        public override NotifyPropertyChanged Clone()
        {
            return new LocalizableRTFString(_value, _value2);
        }
    }
}
