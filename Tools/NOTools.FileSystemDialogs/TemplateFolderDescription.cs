using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    public class TemplateFolderDescription
    {
        public TemplateFolderDescription()
        { 
        }

        public TemplateFolderDescription(string displayName, string path)
        {
            DisplayName = displayName;
            Path = path;
        }


        [DisplayName("Name"), Category("Default"), Description("The shown display name.")]
        public string DisplayName { get; set; }

        [DisplayName("Path"), Category("Default"), Description("The full qualified path.")]
        public string Path { get; set; }

        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        internal TemplateFolderDescriptionCollection Parent { get; set; }

        public override string ToString()
        {
            if (null != DisplayName)
                return DisplayName;
            else
                return "TemplateFolderDescription";
        }
    }
}
