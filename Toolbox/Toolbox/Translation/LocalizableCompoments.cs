using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.Translation
{
    /// <summary>
    /// Localizable Component Collection
    /// </summary>
    internal class LocalizableCompoments : BindingList<LocalizableCompoment>
    {
        /// <summary>
        /// Returns component by name
        /// </summary>
        /// <param name="name">target name</param>
        /// <returns>target component</returns>
        internal LocalizableCompoment this[string name]
        {
            get
            {
                return this.First(c => c.Value.Replace(" ", "") == name.Replace(" ", ""));
            }
        }
    }
}
