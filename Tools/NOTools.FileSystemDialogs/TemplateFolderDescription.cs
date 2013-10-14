using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    /// <summary>
    /// Descripton Item for a custom folder
    /// </summary>
    public class TemplateFolderDescription
    {
        #region Ctor
        
        public TemplateFolderDescription()
        { 
        }

        public TemplateFolderDescription(string displayName, string path)
        {
            DisplayName = displayName;
            Path = path;
        }

        #endregion

        #region Overrides
        
        [DisplayName("Name"), Category("Default"), Description("The shown display name.")]
        public string DisplayName { get; set; }

        [DisplayName("Path"), Category("Default"), Description("The full qualified path.")]
        public string Path { get; set; }

        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        internal TemplateFolderDescriptionCollection Parent { get; set; }
    
        #endregion

        #region Overrides

        public override string ToString()
        {
            if (null != DisplayName)
                return DisplayName;
            else
                return "TemplateFolderDescription";
        }

        #endregion
    }
}
