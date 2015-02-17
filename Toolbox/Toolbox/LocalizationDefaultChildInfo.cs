using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox
{
    /// <summary>
    /// Default ILocalizationChildInfo
    /// </summary>
    internal class LocalizationDefaultChildInfo : ILocalizationChildInfo
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="nameLocalization">caption in language editor</param>
        /// <param name="typeLocalization">type of sub control</param>
        public LocalizationDefaultChildInfo(string nameLocalization, Type typeLocalization)
        {
            NameLocalization = nameLocalization;
            TypeLocalization = typeLocalization;
        }

        public string NameLocalization { get; private set; }

        public Type TypeLocalization { get; private set; }
    }
}
