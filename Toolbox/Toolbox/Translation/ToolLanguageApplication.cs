using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Translation
{
    internal class ToolLanguageApplication
    {
        internal ToolLanguageApplication(ToolLanguage parent)
        {
            Parent = parent;
            Components = new LocalizableCompoments();
            Components.Add(new LocalizableCompoment(Parent, "Error Dialog", typeof(NetOffice.DeveloperToolbox.Controls.Error.ErrorControl)));
            Components.Add(new LocalizableCompoment(Parent, "Language Selector", typeof(TranslationControl)));
        }

        internal ToolLanguage Parent { get; private set; }

        internal LocalizableCompoments Components { get; private set; }
    }
}
