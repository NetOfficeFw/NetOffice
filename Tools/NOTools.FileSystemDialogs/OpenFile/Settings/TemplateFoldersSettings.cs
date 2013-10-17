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
        #region Ctor
        
        internal TemplateFoldersSettings(DefaultSettings defaultSettings, PropertyChangedEventHandler eventHandler = null) : base(defaultSettings, eventHandler)
        {
            FolderTemplates = new TemplateFolderDescriptionCollection();
            DontFireEvents = true;
            PropertyBag.Add("Visible", DefaultBoolean.False);
            DontFireEvents = false;
        }

        #endregion

        #region Properties

        [DisplayName("Custom Folders"), Description("Allows to add custom folders."), DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public TemplateFolderDescriptionCollection FolderTemplates { get; private set; }

        #endregion

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

        public override bool HasAllowedSubFolders(FileSystemInfo fsInfo)
        {
            if (fsInfo is TemplateFolderRoot)
                return true;
            else
                return base.HasAllowedSubFolders(fsInfo);
        }

        public override bool AllowShowFolder(FolderInfo folder)
        {
            return true;
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