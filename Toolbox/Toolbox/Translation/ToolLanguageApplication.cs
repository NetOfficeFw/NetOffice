using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Translation
{
    /// <summary>
    /// Contains all localizable application root components
    /// </summary>
    internal class ToolLanguageApplication
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">the used language for the components</param>
        internal ToolLanguageApplication(ToolLanguage parent)
        {
            Parent = parent;
            Components = new LocalizableCompoments();
            Components.Add(new LocalizableCompoment(Parent, "Error Dialog", typeof(NetOffice.DeveloperToolbox.Controls.Error.ErrorControl)));
            Components.Add(new LocalizableCompoment(Parent, "Language Selector", typeof(TranslationControl)));
        }

        /// <summary>
        /// The used language for the components
        /// </summary>
        internal ToolLanguage Parent { get; private set; }

        /// <summary>
        /// Application components
        /// </summary>
        internal LocalizableCompoments Components { get; private set; }
    }
}
