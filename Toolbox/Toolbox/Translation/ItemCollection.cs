using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.Translation
{
    public class ItemCollection : BindingList<NotifyPropertyChanged>
    {
        public void TryGetValue(string name, out string caption)
        {
            var item = this.Where(i => i.Value.Equals(name, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
            if (null != item)
                caption = item.Value2;
            else
                caption = null;
        }

        public NotifyPropertyChanged this[string value1]
        {
            get
            {
                return this.First(n => n.Value == value1);
            }
        }
    }
}
