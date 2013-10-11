using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    /// <summary>
    /// All TemplateFolders settings for OpenFilePanel.cs
    /// </summary>
    public class TemplateFoldersSettings : DefaultableSettings
    {
        public TemplateFoldersSettings(DefaultSettings defaultSettings, PropertyChangedEventHandler eventHandler = null) : base(defaultSettings, eventHandler)
        {
            DontFireEvents = true;
            PropertyBag.Add("Visible", DefaultBoolean.False);
            DontFireEvents = false;
        }

        #region Overrides

        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public new DefaultBoolean AllowAddFolders
        {
            get { return DefaultBoolean.False; }
            set { }
        }

        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public new DefaultBoolean AllowDeleteFolders
        {
            get { return DefaultBoolean.False; }
            set { }
        }

        /// <summary>
        /// Returns a System.String that represence the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            return "TemplateFolders";
        }

        #endregion
    }
}