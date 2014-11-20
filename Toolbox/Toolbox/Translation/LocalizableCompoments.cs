using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.Translation
{
    internal class LocalizableCompoments : BindingList<LocalizableCompoment>
    {
        internal LocalizableCompoment this[string name]
        {
            get
            {
                return this.First(c => c.Value.Replace(" ", "") == name.Replace(" ", ""));
            }
        }
    }
}
