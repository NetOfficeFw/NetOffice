using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    /// <summary>
    /// All defaultable settings for OpenFilePanel.cs
    /// </summary>
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class DefaultableSettings : INotifyPropertyChanged
    {
        protected internal DefaultableSettings(DefaultSettings defaultSettings, PropertyChangedEventHandler eventHandler = null)
        {
            DefaultSettings = defaultSettings;
            PropertyBag = new PropertyBagCollection<DefaultBoolean>(DefaultBoolean.Default, RaisePropertyChanged);
            if (null != eventHandler)
                this.PropertyChanged += eventHandler;
        }

        #region Properties

        [Category("Defaultable"), Description("Allow user to see category in the left area.")]
        public DefaultBoolean Visible
        {
            get { return PropertyBag["Visible"]; }
            set { PropertyBag["Visible"] = value; }
        }

        [Category("Defaultable"), Description("Get or set the category are expanded.")]
        public DefaultBoolean Expanded
        {
            get { return PropertyBag["Expanded"]; }
            set { PropertyBag["Expanded"] = value; }
        }

        [Category("Defaultable"), Description("Allow user to browse sub directories.")]
        public DefaultBoolean AllowBrowseFolders
        {
            get { return PropertyBag["AllowBrowseFolders"]; }
            set { PropertyBag["AllowBrowseFolders"] = value; }
        }

        [Category("Defaultable"), Description("Allow user to add sub directories.")]
        public DefaultBoolean AllowAddFolders
        {
            get { return PropertyBag["AllowAddFolders"]; }
            set { PropertyBag["AllowAddFolders"] = value; }
        }

        [Category("Defaultable"), Description("Allow user to delete sub directories.")]
        public DefaultBoolean AllowDeleteFolders
        {
            get { return PropertyBag["AllowDeleteFolders"]; }
            set { PropertyBag["AllowDeleteFolders"] = value; }
        }

        [Category("Defaultable"), Description("Allow user delete files")]
        public DefaultBoolean AllowDeleteFiles
        {
            get { return PropertyBag["AllowDeleteFiles"]; }
            set { PropertyBag["AllowDeleteFiles"] = value; }
        }

        [Category("Defaultable"), Description("Allow user to select more than one file.")]
        public DefaultBoolean AllowMultipleSelect
        {
            get { return PropertyBag["AllowMultipleSelect"]; }
            set { PropertyBag["AllowMultipleSelect"] = value; }
        }

        [Category("Defaultable"), Description("Show access errors(may security issues) in a dialog box.")]
        public DefaultBoolean ShowAccessErrorsInDialogBox
        {
            get { return PropertyBag["ShowAccessErrorsInDialogBox"]; }
            set { PropertyBag["ShowAccessErrorsInDialogBox"] = value; }
        }

        /// <summary>
        /// Dynamic property bag to hold property values
        /// </summary>
        protected internal PropertyBagCollection<DefaultBoolean> PropertyBag { get; set; }

        /// <summary>
        /// Flag for silent property changing
        /// </summary>
        protected internal bool DontFireEvents { get; set; }

        /// <summary>
        /// Korresponding default settings
        /// </summary>
        protected internal DefaultSettings DefaultSettings { get; set; }

        #endregion

        #region Methods

        public virtual bool CanCreateFolders(FileSystemInfo fsInfo)
        {
            if (fsInfo is DrvInfoRoot || fsInfo is FiInfo || null == fsInfo)
                return false;
            return true;
        }

        public virtual bool CanDeleteFolders(FileSystemInfo fsInfo)
        {
            if (fsInfo is DrvInfoRoot || fsInfo is FiInfo || null == fsInfo)
                return false;
            return true;
        }

        public virtual bool CanDeleteFiles(FileSystemInfo fsInfo)
        {
            if (fsInfo is DrvInfoRoot || fsInfo is FiInfo || null == fsInfo)
                return false;
            return true;
        }

        public virtual bool HasAllowedSubFolders(FileSystemInfo fsInfo)
        {
            if (!GetRuntimeValue("AllowBrowseFolders"))
                return false;
            
            return fsInfo.HasDirectories;
        }

        public virtual bool AllowShowDrive(DrvInfo drive)
        {
            return true;
        }

        public virtual bool AllowShowFolder(FolderInfo folder)
        {
            return GetRuntimeValue("AllowBrowseFolders");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public virtual bool AllowShowFile(FiInfo file)
        {
            if (file.Name.EndsWith(".lnk", StringComparison.InvariantCultureIgnoreCase))
                return false;

            if (null != DefaultSettings.MiscSettings.CurrentFilter)
            {
                int position = DefaultSettings.MiscSettings.CurrentFilter.Filter.IndexOf(".");
                if (position > 0)
                {
                    string end = DefaultSettings.MiscSettings.CurrentFilter.Filter.Substring(position+1);
                    if (end.Trim() != "*")
                    {
                        if (file.Name.EndsWith(end, StringComparison.InvariantCultureIgnoreCase))
                            return true;
                        else
                            return false;
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Get real value at runtime
        /// </summary>
        /// <param name="propertyName">name of the property</param>
        /// <returns>true or false</returns>
        internal bool GetRuntimeValue(string propertyName)
        {
            DefaultBoolean defBoolValue = PropertyBag[propertyName];
            bool boolVaue = DefaultSettings.PropertyBag[propertyName];
            return IsTrue(boolVaue, defBoolValue);
        }

        private static bool IsTrue(bool defaultBool, DefaultBoolean val)
        {
            if (val == DefaultBoolean.True)
                return true;
            if (val == DefaultBoolean.Default)
                return defaultBool;
            else
                return false;
        
        }

        #endregion

        #region INotifyPropertyChanged

        /// <summary>
        /// Ocurs wenn a property value has changed
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public event PropertyChangedEventHandler PropertyChanged;

        protected internal void RaisePropertyChanged(string propertyName)
        {
            if (null != PropertyChanged && !DontFireEvents)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion
    }
}
