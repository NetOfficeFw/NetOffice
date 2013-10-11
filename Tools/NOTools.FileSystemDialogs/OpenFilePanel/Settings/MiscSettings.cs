using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    /// <summary>
    /// All misc settings for OpenFilePanel.cs
    /// </summary>
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class MiscSettings : INotifyPropertyChanged
    {
        public MiscSettings(PropertyChangedEventHandler eventHandler = null)
        {
            if (null != eventHandler)
                this.PropertyChanged += eventHandler;
        }

        [Category("Misc"), Description("Get or set the category panel in the left area is visible.")]
        public bool ShowCategoryPanel
        {
            get { return _showCategoryPanel; }
            set
            {
                _showCategoryPanel = value;
                RaisePropertyChanged("ShowCategoryPanel");
            }
        }
        private bool _showCategoryPanel = true;

        [Category("Misc"), Description("Get or set the selected category in the left area.")]
        public RootCategory SelectedCategory
        {
            get { return _selectedCategory; }
            set
            {
                _selectedCategory = value;
                RaisePropertyChanged("SelectedCategory");
            }
        }
        private RootCategory _selectedCategory;

        [Category("Misc"), Description("File extenstion filter. identically to the windowsforms openfiledialog.")]
        public string FileFilter
        {
            get { return _fileFilter; }
            set
            {
                _filters = FileFilterItem.CreateFromFilterString(value);
                _fileFilter = value;
                RaisePropertyChanged("FileFilter");
            }
        }
        private string _fileFilter;

        internal FileFilterItem[] Filters
        {
            get
            {
                return _filters;
            }            
        }
        private FileFilterItem[] _filters;

        internal FileFilterItem CurrentFilter { get; set; }
         
        #region INotifyPropertyChanged

        /// <summary>
        /// Ocurs wenn a property value has changed
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public event PropertyChangedEventHandler PropertyChanged;

        private void RaisePropertyChanged(string propertyName)
        {
            if (null != PropertyChanged)
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
            return "Misc";
        }

        #endregion
    }
}
