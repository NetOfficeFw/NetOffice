using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Text;

namespace NOTools.FileSystemDialogs
{  
    /// <summary>
    /// All default settings for OpenFilePanel.cs
    /// </summary>
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class DefaultSettings : INotifyPropertyChanged
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">Parent panel</param>
        internal DefaultSettings(MiscSettings miscSettings, PropertyChangedEventHandler eventHandler = null)
        {
            MiscSettings = miscSettings;
            PropertyBag = new PropertyBagCollection<bool>(false, RaisePropertyChanged, new KeyValuePair<string, bool>[] 
                        { 
                          new KeyValuePair<string, bool>("Visible", true),  
                          new KeyValuePair<string, bool>("AllowBrowseFolders", true), 
                             });
            if (null != eventHandler)
                this.PropertyChanged += eventHandler;
        }

        #endregion

        #region Properties

        [Category("Default"), Description("Expanded state for all categories in the left area.")]
        public bool Expanded
        {
            get { return PropertyBag["Expanded"]; }
            set { PropertyBag["Expanded"] = value; }
        }

        [Category("Default"), Description("Visibility for all categories in the left area.")]
        public bool Visible
        {
            get { return PropertyBag["Visible"]; }
            set { PropertyBag["Visible"] = value; }
        }

        [Category("Default"), Description("Allow user add directories.")]
        public bool AllowAddFolders
        {
            get { return PropertyBag["AllowAddFolders"]; }
            set { PropertyBag["AllowAddFolders"] = value; }
        }

        [Category("Default"), Description("Allow user delete directories.")]
        public bool AllowDeleteFolders
        {
            get { return PropertyBag["AllowDeleteFolders"]; }
            set { PropertyBag["AllowDeleteFolders"] = value; }
        }

        [Category("Default"), Description("Allow user delete files.")]
        public bool AllowDeleteFiles
        {
            get { return PropertyBag["AllowDeleteFiles"]; }
            set { PropertyBag["AllowDeleteFiles"] = value; }
        }

        [Category("Default"), Description("Allow user to browse subdirectories.")]
        public bool AllowBrowseFolders
        {
            get { return PropertyBag["AllowBrowseFolders"]; }
            set { PropertyBag["AllowBrowseFolders"] = value; }
        }

        [Category("Default"), Description("Allow user to select more than one file.")]
        public bool AllowMultipleSelect
        {
            get { return PropertyBag["AllowMultipleSelect"]; }
            set { PropertyBag["AllowMultipleSelect"] = value; }
        }
        
        [Category("Default"), Description("Show access errors(may security issues) in a dialog box.")]
        public bool ShowAccessErrorsInDialogBox
        {
            get { return PropertyBag["ShowAccessErrorsInDialogBox"]; }
            set { PropertyBag["ShowAccessErrorsInDialogBox"] = value; }
        }

        /// <summary>
        /// Korresponding Misc Settings
        /// </summary>
        internal MiscSettings MiscSettings { get; private set; }

        /// <summary>
        /// Dynamic property bag to hold property values
        /// </summary>
        internal PropertyBagCollection<bool> PropertyBag { get; private set; }

        #endregion

        #region INotifyPropertyChanged

        /// <summary>
        /// Ocurs wenn a property value has changed
        /// </summary>
        [EditorBrowsable( EditorBrowsableState.Advanced), Browsable(false)]
        public event PropertyChangedEventHandler PropertyChanged;

        private void RaisePropertyChanged(string propertyName)
        { 
            if(null != PropertyChanged)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion
        
        #region Overrides

        /// <summary>
        /// Returns a System.String that represence the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            return "Default";
        }

        #endregion
    }
}
