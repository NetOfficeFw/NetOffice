using System;

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

        /// <summary>
        /// Caption in language editor
        /// </summary>
        public string NameLocalization { get; private set; }

        /// <summary>
        /// Type of sub control
        /// </summary>
        public Type TypeLocalization { get; private set; }
    }
}
