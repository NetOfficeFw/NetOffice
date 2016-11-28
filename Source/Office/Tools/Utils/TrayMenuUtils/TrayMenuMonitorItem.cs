using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Windows.Forms;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Represents a tray menu diagnostic monitor item
    /// </summary>
    [ItemType(TrayMenuItemType.Monitor)]
    public class TrayMenuMonitorItem : TrayMenuItem
    {
        #region Nested

        /// <summary>
        /// View Options
        /// </summary>
        public class Options : INotifyPropertyChanged
        {         
            private bool _autoExpandNodes = true;
            private bool _highlightNewNodes = true;
            private bool _showConsoleTimeColumn = false;
            private bool _showConsoleKindColumn = false;

            private DataGridView _dataGrid;
            private TreeView _treeView;
            private Timer _timer;
            private Action _disableHighlight;

            internal Options(DataGridView dataGrid, TreeView treeView, Timer timer, Action disableHighlight)
            {
                if (null == dataGrid)
                    throw new ArgumentNullException("dataGrid");
                if (null == treeView)
                    throw new ArgumentNullException("treeView");
                if (null == timer)
                    throw new ArgumentNullException("timer");
                if (null == disableHighlight)
                    throw new ArgumentNullException("disableHighlight");
                _dataGrid = dataGrid;
                _treeView = treeView;
                _timer = timer;
                _disableHighlight = disableHighlight;
            }

            /// <summary>
            /// Auto Expand Proxy Nodes
            /// </summary>
            [DefaultValue(true), Category("Options"), Description("Auto Expand Proxy Nodes")]
            public bool AutoExpandNodes
            {
                get
                {
                    return _autoExpandNodes;
                }
                set
                {
                    if (value != _autoExpandNodes)
                    {
                        _autoExpandNodes = value;
                        if (_autoExpandNodes)
                            _treeView.ExpandAll();
                        RaisePropertyChanged("AutoExpandNodes");
                    }
                }
            }

            /// <summary>
            /// Highlight New Nodes
            /// </summary>
            [DefaultValue(true), Category("Options"), Description("Highlight New Nodes")]
            public bool HighlightNewNodes
            {
                get
                {
                    return _highlightNewNodes;
                }
                set
                {
                    if (value != _highlightNewNodes)
                    {
                        _highlightNewNodes = value;
                        _timer.Enabled = value;
                        RaisePropertyChanged("HighlightNewNodes");
                    }
                }
            }

            /// <summary>
            /// Show Console Time Column
            /// </summary>
            [DefaultValue(true), Category("Options"), Description("Console Time Column Visibility")]
            public bool ShowConsoleTimeColumn
            {
                get
                {
                    return _showConsoleTimeColumn;
                }
                set
                {
                    if (value != _showConsoleTimeColumn)
                    {
                        _showConsoleTimeColumn = value;
                        if (_dataGrid.Columns.Count == 3)
                        {
                            _dataGrid.Columns[1].Visible = value;
                            if (!_dataGrid.Columns[1].Visible || !_dataGrid.Columns[2].Visible)
                                _dataGrid.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                            else
                                _dataGrid.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.NotSet;
                        }
                        RaisePropertyChanged("ShowConsoleTimeColumn");
                    }
                }
            }

            /// <summary>
            /// Show Console Kind Column
            /// </summary>
            [DefaultValue(true), Category("Options"), Description("Console Kind Column Visibility")]
            public bool ShowConsoleKindColumn
            {
                get
                {
                    return _showConsoleKindColumn;
                }
                set
                {
                    if (value != _showConsoleKindColumn)
                    {
                        _showConsoleKindColumn = value;
                        if (_dataGrid.Columns.Count == 3)
                        {
                            _dataGrid.Columns[2].Visible = value;
                            if (!_dataGrid.Columns[1].Visible || !_dataGrid.Columns[2].Visible)
                                _dataGrid.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                            else
                                _dataGrid.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.NotSet;
                        }
                        RaisePropertyChanged("ShowConsoleKindColumn");
                    }
                }
            }

            /// <summary>
            /// Occurs after Property has been changed
            /// </summary>
            public event PropertyChangedEventHandler PropertyChanged;

            private void RaisePropertyChanged(string propertyName)
            {
                if (null != PropertyChanged)
                    PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        /// <summary>
        /// View Category
        /// </summary>
        public enum ShowMode
        {
            /// <summary>
            /// Core Category
            /// </summary>
            Core = 0,

            /// <summary>
            /// Settings Category
            /// </summary>
            Settings = 1,

            /// <summary>
            /// DebugConsole Category
            /// </summary>
            Console = 2,

            /// <summary>
            /// Diagnostics Category
            /// </summary>
            Diagnostics = 3,

            /// <summary>
            /// Proxies Category
            /// </summary>
            Proxies = 4,

            /// <summary>
            /// View Options Category
            /// </summary>
            Options = 5
        }

        #endregion

        #region Fields

        private TrayMenuMonitorItemControl _control;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">item owner</param>
        /// <param name="text">shown caption</param>
        internal TrayMenuMonitorItem(TrayMenu owner, string text) : base(owner, text)
        {
            ItemType = TrayMenuItemType.Monitor;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">item owner</param>
        /// <param name="text">shown caption</param>
        /// <param name="visible">item visibility</param>
        internal TrayMenuMonitorItem(TrayMenu owner, string text, bool visible) : base(owner, text, visible)
        {
            ItemType = TrayMenuItemType.Monitor;
        }

        #endregion

        #region Properties

        /// <summary>
        /// View options in user interface
        /// </summary>
        public Options ViewOptions
        {
            get
            {
                return _control.ViewOptions;
            }
            set
            {
                _control.ViewOptions = value;
            }
        }

        /// <summary>
        /// Current shown category
        /// </summary>
        public ShowMode Mode
        {
            get
            {
                return _control.GetCurrentMode();            
            }
            set
            {
                _control.SetCurrentMode(value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Enable or disable category visibilty. Core is ignored
        /// </summary>
        /// <param name="mode">mode to set</param>
        /// <param name="visible">visibility</param>
        public void SetModeVisible(TrayMenuMonitorItem.ShowMode mode, bool visible)
        {
            _control.SetModeVisible(mode, visible);
        }

        /// <summary>
        /// Get current category visibilty
        /// </summary>
        /// <param name="mode">target mode</param>
        /// <returns>visibility</returns>
        public bool GetModeVisible(TrayMenuMonitorItem.ShowMode mode)
        {
            return _control.GetModeVisible(mode);
        }

        internal void SetMonitorElements(TrayMenuMonitorItemControl control)
        {
            if (null == control)
                throw new ArgumentNullException("control");
            _control = control;
        }

        #endregion
    }
}