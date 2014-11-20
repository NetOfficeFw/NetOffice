using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox
{
    internal class LocalizationDefaultChildInfo : ILocalizationChildInfo
    {
        public LocalizationDefaultChildInfo(string nameLocalization, Type typeLocalization)
        {
            NameLocalization = nameLocalization;
            TypeLocalization = typeLocalization;
        }

        public string NameLocalization { get; private set; }

        public Type TypeLocalization { get; private set; }
    }
}
