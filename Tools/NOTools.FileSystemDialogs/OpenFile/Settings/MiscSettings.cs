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
                this.InternalPropertyChanged += eventHandler;
        }

        [Category("Misc"), Description("Get or set a message box appears and askt the user to confirm a delete action.")]
        public bool AskBeforeDelete
        {
            get { return _askBeforeDelete; }
            set
            {
                _askBeforeDelete = value;
                RaiseInternalPropertyChanged("AskBeforeDelete");
                RaisePropertyChanged("AskBeforeDelete");
            }
        }
        private bool _askBeforeDelete = true;

        [Category("Misc"), Description("Get or set the file selection panel in the bottom area is visible.")]
        public bool ShowFilePanel
        {
            get { return _showFilePanel; }
            set
            {
                _showFilePanel = value;
                RaiseInternalPropertyChanged("ShowFilePanel");
                RaisePropertyChanged("ShowFilePanel");
            }
        }
        private bool _showFilePanel = true;

        [DisplayName("NoDoubleClickEvent"), Category("Misc"), Description("The control fires the SelectionChanged event instead of FileDoubleClick.")]
        public bool FireSelectionChangedInsteadOfDoubleClick
        {
            get { return _fireSelectionChangedInsteadOfDoubleClick; }
            set
            {
                _fireSelectionChangedInsteadOfDoubleClick = value;
                RaiseInternalPropertyChanged("FireSelectionChangedInsteadOfDoubleClick");
                RaisePropertyChanged("FireSelectionChangedInsteadOfDoubleClick");
            }
        }
        private bool _fireSelectionChangedInsteadOfDoubleClick;

        [Category("Misc"), Description("Get or set the category panel in the left area is visible.")]
        public bool ShowCategoryPanel
        {
            get { return _showCategoryPanel; }
            set
            {
                _showCategoryPanel = value;
                RaiseInternalPropertyChanged("ShowCategoryPanel");
                RaisePropertyChanged("ShowCategoryPanel");
            }
        }
        private bool _showCategoryPanel = true;

        [Category("Misc"), Description("Get or set the selected category in the left area.")]
        [RefreshProperties(RefreshProperties.All)]
        public RootCategory SelectedCategory
        {
            get { return _selectedCategory; }
            set
            {
                _selectedCategory = value;
                _selectedFiles = new string[0];
                _selectedFile = string.Empty;
                RaiseInternalPropertyChanged("SelectedCategory");
                RaisePropertyChanged("SelectedCategory");
            }
        }
        private RootCategory _selectedCategory;

        [Category("Misc"), Description("Current selected file.")]
        [RefreshProperties(RefreshProperties.All)]
        public string SelectedFile
        {
            get { return _selectedFile; }
        }
        private string _selectedFile;

        [Category("Misc"), Description("Current selected files. if multi select enabled.")]
        [RefreshProperties(RefreshProperties.All)]
        public string[] SelectedFiles
        {
            get { return _selectedFiles; }
        }
        private string[] _selectedFiles = new string[0];

        [Category("Misc"), Description("Width of category panel in the left area.")]
        [RefreshProperties(RefreshProperties.All)]
        public int CategoryPanelWidth
        {
            get { return _categoryPanelWidth; }
            set
            {
                RaiseInternalPropertyChanged("CategoryPanelWidth");
                RaisePropertyChanged("CategoryPanelWidth");
                _categoryPanelWidth = value;
            }
        }
        private int _categoryPanelWidth;


        [Category("Misc"), Description("File extenstion filter. identically to the windowsforms openfiledialog.")]
        [RefreshProperties(RefreshProperties.All)]
        public string FileFilter
        {
            get { return _fileFilter; }
            set
            {
                _filters = FileFilterItem.CreateFromFilterString(value);
                _fileFilter = value;
                RaiseInternalPropertyChanged("FileFilter");
                RaisePropertyChanged("FileFilter");
            }
        }
        private string _fileFilter = string.Empty;

        internal FileFilterItem[] Filters
        {
            get
            {
                return _filters;
            }            
        }
        private FileFilterItem[] _filters;

        internal FileFilterItem CurrentFilter { get; set; }

        internal void SetSelectedCategory(RootCategory category)
        {
            _selectedCategory = category;

            _selectedFiles = new string[0];
            _selectedFile = string.Empty;
            RaisePropertyChanged("SelectedCategory");
            RaisePropertyChanged("SelectedDirectory");
            RaisePropertyChanged("SelectedFiles");
            RaisePropertyChanged("SelectedFile");
        }

        internal void SetCategoryPanelWidth(int width)
        {
            _categoryPanelWidth = width;
            RaisePropertyChanged("CategoryPanelWidth");
        }

        internal void SetSelectedFile(string file)
        {
            _selectedFile = file;
            _selectedFiles = new string[0];
            RaisePropertyChanged("SelectedFiles");
            RaisePropertyChanged("SelectedFile");
        }

        internal void SetSelectedFiles(string[] files)
        {
            _selectedFiles = files;
            if (null != _selectedFiles && _selectedFiles.Length > 0)
                _selectedFile = files[0];
            else
                _selectedFile = String.Empty;
            RaisePropertyChanged("SelectedFiles");
            RaisePropertyChanged("SelectedFile");
        }

        #region INotifyPropertyChanged

        /// <summary>
        /// Ocurs wenn a property value has changed
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false)]
        public event PropertyChangedEventHandler InternalPropertyChanged;

        private void RaiseInternalPropertyChanged(string propertyName)
        {
            if (null != InternalPropertyChanged)
                InternalPropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

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
