using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.Translation
{
    /// <summary>
    /// Localizable Resource Item Collection
    /// </summary>
    public class ItemCollection : BindingList<NotifyPropertyChanged>
    {
        /// <summary>
        /// Try get a localizable string value
        /// </summary>
        /// <param name="name">name of the value</param>
        /// <param name="caption">target value</param>
        public void TryGetValue(string name, out string caption)
        {
            var item = this.Where(i => i.Value.Equals(name, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
            if (null != item)
                caption = item.Value2;
            else
                caption = null;
        }

        /// <summary>
        /// Get Resource item by name
        /// </summary>
        /// <param name="value1">target name</param>
        /// <returns>target value</returns>
        public NotifyPropertyChanged this[string value1]
        {
            get
            {
                return this.First(n => n.Value == value1);
            }
        }
    }
}
