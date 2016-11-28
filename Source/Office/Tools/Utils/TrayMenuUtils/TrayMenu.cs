using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Reflection;
using System.Text;
using System.Runtime;
using System.Collections;
using NetOffice.Tools;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Represents an optional overlay item menu for a tray icon
    /// </summary>
    public class TrayMenu : IDisposableState
    {
        #region Fields

        private ContextMenuStrip _contextMenu;
        private bool _enabled;
        private COMAddinBase _addin;

        #endregion
        
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>       
        internal TrayMenu(COMAddinBase addin)
        {
            _addin = addin;
            _contextMenu = new ContextMenuStrip();
            _contextMenu.Font = new System.Drawing.Font("Arial", 8.00F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(161))); 
            _contextMenu.Opening += ContextMenu_Opening;
            Items = new TrayMenuItems(this);
            Enabled = true;
        }

        #endregion

        #region Events
        
        /// <summary>
        /// Occurs when a selected object from TrayMenuDropDownList has been changed
        /// </summary>
        public event TrayMenuItemSelectedObjectChangedEventHandler ItemSelectedObjectChanged;

        /// <summary>
        /// Occurs when an item text has been changed
        /// </summary>
        public event TrayMenuItemTextChangedEventHandler ItemTextChanged;

        /// <summary>
        /// Occurs when a key is pressed while the control has focus. 
        /// </summary>
        public event TrayMenuItemKeyEventHandler ItemKeyDown;

        /// <summary>
        /// Occurs when a key is pressed while the control has focus.
        /// </summary>
        public event TrayMenuItemKeyPressEventHandler ItemKeyPress;

        /// <summary>
        /// Occurs when a key is released while the control has focus.
        /// </summary>
        public event TrayMenuItemKeyEventHandler ItemKeyUp;

        /// <summary>
        /// Occurs when an item has been clicked
        /// </summary>
        public event TrayMenuItemClickEventHandler ItemClick;

        /// <summary>
        /// Occurs when an item checked state has been changed
        /// </summary>
        public event TrayMenuItemCheckedEventHandler ItemChecked;

        /// <summary>
        /// Occurs before menu is shown
        /// </summary>
        public event TrayMenuOpeningHandlerEventHandler Opening;

        #endregion

        #region Properties
      
        /// <summary>
        /// Addin Owner
        /// </summary>
        internal COMAddinBase Owner
        {
            get
            {
                return _addin;
            }
        }

        /// <summary>
        /// Close on outer click or lost focus
        /// </summary>
        public bool AutoClose
        {
            get
            {
                return _contextMenu.AutoClose;
            }
            set
            {
                if (value != _contextMenu.AutoClose)
                {
                    _contextMenu.AutoClose = value;
                    OnAutoCloseChanged();
                }
            }
        }

        /// <summary>
        /// Contains all items in the tray icon menu
        /// </summary>
        public TrayMenuItems Items { get; private set; }

        /// <summary>
        /// Mouse Button Mode
        /// </summary>
        public TrayMenuClickMode ClickMode { get; private set; }

        /// <summary>
        /// Enable or disable menu
        /// </summary>
        public bool Enabled
        {
            get
            {
                return _enabled;
            }
            set
            {
                if (value != _enabled)
                {
                    _contextMenu.Visible = value;
                    _enabled = value;
                }
            }
        }

        /// <summary>
        /// Item Image Size
        /// </summary>
        public Size ImageScalingSize
        {
            get
            {
                return _contextMenu.ImageScalingSize;
            }
            set
            {
                _contextMenu.ImageScalingSize = value;
            }
        }

        /// <summary>
        /// Show checkbox state on the left
        /// </summary>
        public bool ShowCheckMargin
        {
            get
            {
                return _contextMenu.ShowCheckMargin;
            }
            set
            {
                _contextMenu.ShowCheckMargin = value;
            }
        }

        /// <summary>
        /// Show menu item images on the left
        /// </summary>
        public bool ShowImageMargin
        {
            get
            {
                return _contextMenu.ShowImageMargin;
            }
            set
            {
                _contextMenu.ShowImageMargin = value;
            }
        }

        /// <summary>
        /// Menu background color
        /// </summary>
        public Color BackColor
        {
            get
            {
                return _contextMenu.BackColor;
            }
            set
            {
                _contextMenu.BackColor = value;
            }
        }

        /// <summary>
        /// Menu Default font
        /// </summary>
        public Font Font
        {
            get
            {
                return _contextMenu.Font;
            }
            set
            {
                _contextMenu.Font = value;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Close the menu if open
        /// </summary>
        public void Close()
        {
            _contextMenu.Close();
        }

        /// <summary>
        /// called when AutoClose has been changed
        /// </summary>
        internal void OnAutoCloseChanged()
        {
            // recursive notify auto close items
        }

        /// <summary>
        /// Returns the internal representation from current used ui system
        /// </summary>
        /// <returns>inner ui instance or null</returns>
        internal object GetMenuInternal()
        {
            return _contextMenu;
        }

        /// <summary>
        /// Returns the internal representation from current used ui system
        /// </summary>
        /// <returns>inner ui instance or null</returns>
        internal T GetMenuInternal<T>() where T: class
        {
            return _contextMenu as T;
        }

        /// <summary>
        /// Raise the ItemClick event
        /// </summary>
        /// <param name="item">target item</param>
        internal void RaiseItemClick(TrayMenuItem item)
        {
            if (null != ItemClick)
                ItemClick(this, new TrayMenuItemsEventArgs(item));
        }

        /// <summary>
        /// Raise the ItemChecked event
        /// </summary>
        /// <param name="item">target item</param>
        /// <param name="check">target item checked state</param>
        private void RaiseItemChecked(TrayMenuItem item, bool check)
        {
            if (null != ItemChecked)
                ItemChecked(this, new TrayMenuItemCheckedEventArgs(item, check));
        }

        /// <summary>
        /// Raise the ItemTextChanged event
        /// </summary>
        /// <param name="item">target item</param>
        private void RaiseItemTextChanged(TrayMenuItem item)
        {
            if (null != ItemTextChanged)
                ItemTextChanged(this, new TrayMenuItemTextChangedEventArgs(item));
        }

        /// <summary>
        /// Raise the ItemKeyDown event
        /// </summary>
        /// <param name="item">target item</param>
        /// <param name="args">event arguments</param>
        private void RaiseItemKeyDown(TrayMenuItem item, KeyEventArgs args)
        {
            if (null != ItemKeyDown)
            {
                ToolsKeys keyData = (ToolsKeys)args.KeyData;
                ToolsKeys keyCode = (ToolsKeys)args.KeyCode;
                ToolsKeys modifiers = (ToolsKeys)args.Modifiers;
                TrayMenuItemKeyEventArgs wrapperArgs = new TrayMenuItemKeyEventArgs(item, keyData, args.Alt, args.Control, args.Handled, keyCode, args.KeyValue, modifiers, args.Shift, args.SuppressKeyPress);
                ItemKeyDown(this, wrapperArgs);
                args.Handled = wrapperArgs.Handled;
            }
        }

        /// <summary>
        /// Raise the ItemKeyPress event
        /// </summary>
        /// <param name="item">target item</param>
        /// <param name="args">event arguments</param>
        private void RaiseItemKeyPress(TrayMenuItem item, KeyPressEventArgs args)
        {
            if (null != ItemKeyPress)
            {
                TrayMenuItemKeyPressEventArgs wrapperArgs = new TrayMenuItemKeyPressEventArgs(args.KeyChar);
                ItemKeyPress(this, wrapperArgs);
                args.Handled = wrapperArgs.Handled;
            }
        }

        /// <summary>
        /// Raise the ItemKeyUp event
        /// </summary>
        /// <param name="item">target item</param>
        /// <param name="args">event arguments</param>
        private void RaiseItemKeyUp(TrayMenuItem item, KeyEventArgs args)
        {
            if (null != ItemKeyUp)
            {
                ToolsKeys keyData = (ToolsKeys)args.KeyData;
                ToolsKeys keyCode = (ToolsKeys)args.KeyCode;
                ToolsKeys modifiers = (ToolsKeys)args.Modifiers;
                TrayMenuItemKeyEventArgs wrapperArgs = new TrayMenuItemKeyEventArgs(item, keyData, args.Alt, args.Control, args.Handled, keyCode, args.KeyValue, modifiers, args.Shift, args.SuppressKeyPress);
                ItemKeyUp(this, wrapperArgs);
                args.Handled = wrapperArgs.Handled;
            }
        }

        /// <summary>
        /// Raise the ItemSelectedObjectChanged event
        /// </summary>
        /// <param name="item">target item</param>
        /// <param name="selectedObject">current selected object</param>
        /// <param name="selectedIndex">current selected index</param>
        private void RaiseItemSelectedObjectChanged(TrayMenuItem item, object selectedObject, int selectedIndex)
        {
            if (null != ItemSelectedObjectChanged)
                ItemSelectedObjectChanged(this, new TrayMenuItemSelectedObjectChangedEventArgs(item, selectedObject, selectedIndex));
        }

        /// <summary>
        /// Notify an item visibility has been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal void OnItemVisibleChanged(TrayMenuItem item)
        {
            ToolStripItem targetStrip = Find(item);
            if (null != targetStrip)
                targetStrip.Visible = item.Visible;   
        }

        /// <summary>
        /// Notify an item text has been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal void OnItemTextChanged(TrayMenuItem item)
        {
            ToolStripItem targetStrip = Find(item);
            if (null != targetStrip)
                targetStrip.Text = item.Text;

            RaiseItemTextChanged(item);
        }
        
        /// <summary>
        /// Notify an item enabled state has been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal void OnItemEnabledChanged(TrayMenuItem item)
        {
            ToolStripItem targetStrip = Find(item);
            if (null != targetStrip)
                targetStrip.Enabled = item.Enabled;
        }

        /// <summary>
        /// Notify an item tooltip has been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal void OnItemToolTipTextChanged(TrayMenuItem item)
        {
            ToolStripItem targetStrip = Find(item);
            if (null != targetStrip)
                targetStrip.ToolTipText = item.ToolTipText;
        }

        /// <summary>
        /// Notify an item image has been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal void OnItemImageChanged(TrayMenuItem item)
        {
            ToolStripItem targetStrip = Find(item);
            if (null != targetStrip)
                targetStrip.Image = item.Image;
        }

        /// <summary>
        /// Notify an item back color has been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal void OnItemBackColorChanged(TrayMenuItem item)
        {
            ToolStripItem targetStrip = Find(item);
            if (null != targetStrip)
                targetStrip.BackColor = item.BackColor;
        }

        /// <summary>
        /// Notify an item fore color has been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal void OnItemForeColorChanged(TrayMenuItem item)
        {
            ToolStripItem targetStrip = Find(item);
            if (null != targetStrip)
                targetStrip.ForeColor = item.ForeColor;
        }

        /// <summary>
        /// Notify an item font has been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal void OnItemFontChanged(TrayMenuItem item)
        {
            ToolStripItem targetStrip = Find(item);
            if (null != targetStrip)
                targetStrip.Font = item.Font;
        }

        /// <summary>
        /// Notify an item text align has been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal void OnItemTextAlignChanged(TrayMenuItem item)
        {
            ToolStripItem targetStrip = Find(item);
            if (null != targetStrip)
                targetStrip.TextAlign = item.TextAlign;
        }
        
        /// <summary>
        /// Notify an item padding has been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal void OnItemPaddingChanged(TrayMenuItem item)
        {
            ToolStripItem targetStrip = Find(item);
            if (null != targetStrip)
                targetStrip.Padding = item.Padding;
        }

        /// <summary>
        /// Notify an item image align has been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal void OnItemImageAlignChanged(TrayMenuItem item)
        {
            ToolStripItem targetStrip = Find(item);
            if (null != targetStrip)
                targetStrip.ImageAlign = item.ImageAlign;
        }

        /// <summary>
        /// Notify a list item data source has new item
        /// </summary>
        /// <param name="item">target item</param>
        /// <param name="listItem">new added list item</param>
        internal void OnDropDownItem_ListItemAdded(TrayMenuDropDownListItem item, object listItem)
        {
            ToolStripComboBox targetStrip = Find(item) as ToolStripComboBox;
            if (null != targetStrip)
                targetStrip.Items.Add(listItem);
        }

        /// <summary>
        /// Notify a list item data source has a removed item
        /// </summary>
        /// <param name="item">target item</param>
        /// <param name="listItem">removed item</param>
        /// <param name="listItemIndex">removed item former item index</param>
        internal void OnDropDownItem_ListItemRemoved(TrayMenuDropDownListItem item, object listItem, int listItemIndex)
        {
            ToolStripComboBox targetStrip = Find(item) as ToolStripComboBox;
            if (null != targetStrip)
                targetStrip.Items.Remove(listItem);
        }

        /// <summary>
        /// Notify a list item data source has been cleared
        /// </summary>
        /// <param name="item">target item</param>
        internal void OnDropDownItem_ListItemsCleared(TrayMenuDropDownListItem item)
        {
            ToolStripComboBox targetStrip = Find(item) as ToolStripComboBox;
            if (null != targetStrip)
                targetStrip.Items.Clear();
        }

        /// <summary>
        /// Notify an item dropdown style has been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal void OnDropDownItemStyleChanged(TrayMenuDropDownListItem item)
        {
            ToolStripComboBox targetStrip = Find(item) as ToolStripComboBox;
            if (null != targetStrip)
            {
                targetStrip.DropDownStyle = (ComboBoxStyle)item.DropDownStyle;                   
            }
        }

        /// <summary>
        /// Notify an item dropdown height has been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal int OnDropDownItemDropDownHeightChanged(TrayMenuDropDownListItem item)
        {
            ToolStripComboBox targetStrip = Find(item) as ToolStripComboBox;
            if (null != targetStrip)
            {
                try
                {
                    targetStrip.DropDownHeight = item.DropDownHeight;
                }
                catch
                {
                    ;
                }
                return targetStrip.DropDownHeight;
            }
            else
                return item.DropDownHeight;
        }

        /// <summary>
        /// Notify an item dropdown max length has been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal int OnDropDownItemMaxLengthChanged(TrayMenuDropDownListItem item)
        {
            ToolStripComboBox targetStrip = Find(item) as ToolStripComboBox;
            if (null != targetStrip)
            {
                try
                {
                    targetStrip.MaxLength = item.MaxLength;
                }
                catch
                {
                    ;
                }
                return targetStrip.MaxLength;
            }
            else
                return item.MaxLength;
        }

        /// <summary>
        /// Notify an item text max length has been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal int OnTextBoxItemMaxLengthChanged(TrayMenuTextboxItem item)
        {
            ToolStripTextBox targetStrip = Find(item) as ToolStripTextBox;
            if (null != targetStrip)
            {
                try
                {
                    targetStrip.MaxLength = item.MaxLength;
                }
                catch
                {
                    ;
                }
                return targetStrip.MaxLength;
            }
            else
                return item.MaxLength;
        }       
    
        /// <summary>
        /// Notify an item check state been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal void OnItemCheckedChanged(TrayMenuCheckboxItem item)
        {
            ToolStripItem targetStrip = Find(item);
            if (null != targetStrip)
            {
                ToolStripMenuItem menuItem = targetStrip as ToolStripMenuItem;
                if (null != menuItem)
                    menuItem.Checked = item.Checked;
            }
        }

        /// <summary>
        /// Notify item progress minimum has been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal int OnProgressItemMinimumChanged(TrayMenuProgressItem item)
        {
            ToolStripProgressBar targetStrip = Find(item) as ToolStripProgressBar;
            if (null != targetStrip)
            {
                try
                {
                    targetStrip.Minimum = item.Minimum;
                }
                catch
                {
                    ;
                }
            }
            return targetStrip.Minimum;
        }

        /// <summary>
        /// Notify item progress maximum has been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal int OnProgressItemMaximumChanged(TrayMenuProgressItem item)
        {
            ToolStripProgressBar targetStrip = Find(item) as ToolStripProgressBar;
            if (null != targetStrip)
            {
                try
                {
                    targetStrip.Maximum = item.Maximum;
                }
                catch
                {
                    ;
                }
            }
            return targetStrip.Maximum;
        }

        /// <summary>
        /// Notify item progress value has been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal int OnProgressItemValueChanged(TrayMenuProgressItem item)
        {
            ToolStripProgressBar targetStrip = Find(item) as ToolStripProgressBar;
            if (null != targetStrip)
            {
                try
                {
                    targetStrip.Value = item.Value;
                }
                catch
                {
                    ;
                }
            }
            return targetStrip.Value;
        }

        /// <summary>
        /// Notify item progress style has been changed
        /// </summary>
        /// <param name="item">target item</param>
        internal void OnProgressItemStyleChanged(TrayMenuProgressItem item)
        {
            ToolStripProgressBar targetStrip = Find(item) as ToolStripProgressBar;
            if (null != targetStrip)
            {
                targetStrip.Style = (ProgressBarStyle)item.Style;
            }
        }

        /// <summary>
        /// Notify a new item has been added
        /// </summary>
        /// <param name="parentItem">parent item instance or null</param>
        /// <param name="item">new created item</param>
        internal void OnItemAdded(TrayMenuItem parentItem, TrayMenuItem item)
        {
            ToolStripItem newItem = null;
            switch (item.ItemType)
            {
                case TrayMenuItemType.Item:
                    newItem = new ToolStripMenuItem();                   
                    break;
                case TrayMenuItemType.Label:
                    newItem = new ToolStripLabel();
                    break;
                case TrayMenuItemType.LinkLabel:
                    newItem = new ToolStripLabel();
                    (newItem as ToolStripLabel).IsLink = true;
                    (newItem as ToolStripLabel).LinkBehavior = LinkBehavior.AlwaysUnderline;
                    break;
                case TrayMenuItemType.Button:
                    newItem = new ToolStripButton();
                    break;
                case TrayMenuItemType.TextBox:
                    newItem = new ToolStripTextBox();
                    (newItem as ToolStripTextBox).TextChanged += ToolStripTextBox_TextChanged;
                    break;
                case TrayMenuItemType.CheckBox:
                    newItem = new ToolStripMenuItem();
                    (newItem as ToolStripMenuItem).CheckOnClick = true;
                    (newItem as ToolStripMenuItem).CheckedChanged += ToolStripItem_CheckedChanged;
                    break;
                case TrayMenuItemType.Progress:
                    newItem = new ToolStripProgressBar();
                    (item as TrayMenuProgressItem).SetProgressElements((newItem as ToolStripProgressBar).Minimum, (newItem as ToolStripProgressBar).Maximum, (newItem as ToolStripProgressBar).Value, (TrayMenuProgressItem.ProgressBarStyle)(newItem as ToolStripProgressBar).Style);                    
                    break;
                case TrayMenuItemType.DropDownList:
                    newItem = new ToolStripComboBox();
                    (newItem as ToolStripComboBox).DropDownStyle = ComboBoxStyle.DropDownList;
                    (newItem as ToolStripComboBox).TextChanged += ToolStripComboBox_TextChanged;
                    (newItem as ToolStripComboBox).SelectedIndexChanged += ToolStripComboBox_SelectedIndexChanged;
                    (item as TrayMenuDropDownListItem).SetupDropDownElements((newItem as ToolStripComboBox).DropDownHeight);
                    break;
                case TrayMenuItemType.Separator:
                    newItem = new ToolStripSeparator();
                    break;
                case TrayMenuItemType.Custom:          
                    newItem = new ToolStripControlHost((System.Windows.Forms.Control)(item as TrayMenuCustomItem).Control);
                    break;
                case TrayMenuItemType.Monitor:
                    TrayMenuMonitorItemControl control = new TrayMenuMonitorItemControl(_addin);
                    newItem = new ToolStripControlHost(control);
                    (item as TrayMenuMonitorItem).SetMonitorElements(control);
                    break;
                case TrayMenuItemType.AutoClose:
                    newItem = new ToolStripMenuItem();
                    (newItem as ToolStripMenuItem).CheckOnClick = true;
                    (newItem as ToolStripMenuItem).Checked = AutoClose;                
                    (newItem as ToolStripMenuItem).CheckedChanged += ToolAutoCloseStripItem_CheckedChanged;
                    break;
                case TrayMenuItemType.Close:
                    newItem = new ToolStripMenuItem();
                    (newItem as ToolStripMenuItem).Click += ToolCloseStripItem_Click;
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
            
            newItem.Tag = item;
            newItem.Text = item.Text;
            newItem.ToolTipText = item.ToolTipText;
            newItem.Image = item.Image;
            newItem.BackColor = item.BackColor;
            newItem.ForeColor = item.ForeColor;
            newItem.Visible = item.Visible;
            newItem.Padding = item.Padding;
            newItem.Enabled = item.Enabled;
            item.SetupElements(newItem.Font, newItem.TextAlign, newItem.ImageAlign, item.Padding);

            ToolStripControlHost stripHost = newItem as ToolStripControlHost;
            if (null != stripHost)
            {
                stripHost.KeyDown += StripHost_KeyDown;
                stripHost.KeyPress += StripHost_KeyPress;
                stripHost.KeyUp += StripHost_KeyUp;
            }

            if (null != parentItem)
            {
                ToolStripDropDownItem parentStrip = Find(parentItem) as ToolStripDropDownItem;
                if (null != parentStrip)
                    parentStrip.DropDownItems.Add(newItem);
            }
            else
            {
                _contextMenu.Items.Add(newItem);
            }

            newItem.Click += ToolStripItem_Click;
        }

        /// <summary>
        /// Notify an item has been removed
        /// </summary>
        /// <param name="parentItem">parent item</param>
        /// <param name="item">removed item</param>
        /// <param name="itemIndex">former item index</param>
        internal void OnItemRemoved(TrayMenuItem parentItem, TrayMenuItem item, int itemIndex)
        {
            if (null != parentItem)
            {
                ToolStripItem parentStrip = Find(parentItem);
                if (null != parentStrip)
                {
                    ToolStripItem toolStripItem = Find(item);
                    toolStripItem.Click -= ToolStripItem_Click;
                    _contextMenu.Items.Remove(toolStripItem);
                }
            }
            else
            {
                ToolStripItem toolStripItem = Find(item);
                toolStripItem.Click -= ToolStripItem_Click;
                _contextMenu.Items.Remove(toolStripItem);
            }
        }

        /// <summary>
        /// Notify all item has been removed
        /// </summary>
        internal void OnItemsClear()
        {
            _contextMenu.Items.Clear();
        }

        /// <summary>
        /// Find an item in the ui collection
        /// </summary>
        /// <param name="item">target item</param>
        /// <returns>ui item or null</returns>
        private ToolStripItem Find(TrayMenuItem item)
        {
            foreach (ToolStripItem stripItem in _contextMenu.Items)
            {
                TrayMenuItem trayItem = stripItem.Tag as TrayMenuItem;
                if (trayItem == item)
                    return stripItem;

                ToolStripDropDownItem dropDownItem = stripItem as ToolStripDropDownItem;
                if (null != dropDownItem)
                {
                    ToolStripItem result = Find(item, dropDownItem.DropDownItems);
                    if (null != result)
                        return result;
                }
            }
            return null;
        }

        /// <summary>
        /// Recursive find item helper
        /// </summary>
        /// <param name="item">target item</param>
        /// <param name="items">collection to search</param>
        /// <returns>item or null</returns>
        private ToolStripItem Find(TrayMenuItem item, ToolStripItemCollection items)
        {
            foreach (ToolStripItem stripItem in items)
            {
                TrayMenuItem trayItem = stripItem.Tag as TrayMenuItem;
                if (trayItem == item)
                    return stripItem;
                ToolStripDropDownItem dropDownItem = stripItem as ToolStripDropDownItem;
                if (null != dropDownItem)
                {
                    ToolStripItem result = Find(item, dropDownItem.DropDownItems);
                    if (null != result)
                        return result;
                }
            }
            return null;
        }

        #endregion

        #region IDisposableState

        /// <summary>
        /// Returns information the instance is already disposed
        /// </summary>
        public bool IsDisposed { get; private set; }      

        /// <summary>
        /// Dispose the instance
        /// </summary>
        public void Dispose()
        {
            if (IsDisposed)
            {
                if (null != _contextMenu)
                { 
                    _contextMenu.Dispose();
                    _contextMenu = null;
                }
                IsDisposed = true;
            }
        }

        #endregion


        #region Trigger

        private void ToolCloseStripItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void ContextMenu_Opening(object sender, CancelEventArgs e)
        {
            try
            {
                if (!Enabled)
                    e.Cancel = true;

                if (null != Opening)
                {
                    CancelEventArgs args = new CancelEventArgs();
                    Opening(this, args);
                    e.Cancel = args.Cancel;
                }
            }
            catch (Exception exception)
            {
                NetOffice.Core.Default.Console.WriteException(exception);
            }
        }

        private void ToolStripItem_Click(object sender, EventArgs e)
        {
            try
            {
                ToolStripItem toolStripItem = sender as ToolStripItem;
                TrayMenuItem menuItem = toolStripItem.Tag as TrayMenuItem;
                RaiseItemClick(menuItem);
            }
            catch (Exception exception)
            {
                NetOffice.Core.Default.Console.WriteException(exception);
            }
        }

        private void ToolAutoCloseStripItem_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                ToolStripMenuItem toolStripItem = sender as ToolStripMenuItem;
                AutoClose = toolStripItem.Checked;
            }
            catch (Exception exception)
            {
                NetOffice.Core.Default.Console.WriteException(exception);
            }
        }

        private void ToolStripItem_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                ToolStripMenuItem toolStripItem = sender as ToolStripMenuItem;
                TrayMenuItem menuItem = toolStripItem.Tag as TrayMenuItem;
                RaiseItemChecked(menuItem, toolStripItem.Checked);             
            }
            catch (Exception exception)
            {
                NetOffice.Core.Default.Console.WriteException(exception);
            }
        }

        private void ToolStripTextBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                ToolStripTextBox toolStripItem = sender as ToolStripTextBox;
                TrayMenuItem menuItem = toolStripItem.Tag as TrayMenuItem;
                RaiseItemTextChanged(menuItem);
            }
            catch (Exception exception)
            {
                NetOffice.Core.Default.Console.WriteException(exception);
            }
        }

        private void ToolStripComboBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                ToolStripComboBox toolStripItem = sender as ToolStripComboBox;
                TrayMenuItem menuItem = toolStripItem.Tag as TrayMenuItem;
                RaiseItemTextChanged(menuItem);
            }
            catch (Exception exception)
            {
                NetOffice.Core.Default.Console.WriteException(exception);
            }
        }

        private void ToolStripComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                ToolStripComboBox toolStripItem = sender as ToolStripComboBox;
                TrayMenuItem menuItem = toolStripItem.Tag as TrayMenuItem;
                RaiseItemSelectedObjectChanged(menuItem, toolStripItem.SelectedItem, toolStripItem.SelectedIndex);
            }
            catch (Exception exception)
            {
                NetOffice.Core.Default.Console.WriteException(exception);
            }
        }

        private void StripHost_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                ToolStripControlHost toolStripItem = sender as ToolStripControlHost;
                TrayMenuItem menuItem = toolStripItem.Tag as TrayMenuItem;
                RaiseItemKeyDown(menuItem, e);   
            }
            catch (Exception exception)
            {
                NetOffice.Core.Default.Console.WriteException(exception);
            }
        }

        private void StripHost_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                ToolStripControlHost toolStripItem = sender as ToolStripControlHost;
                TrayMenuItem menuItem = toolStripItem.Tag as TrayMenuItem;
                RaiseItemKeyUp(menuItem, e);               
            }
            catch (Exception exception)
            {
                NetOffice.Core.Default.Console.WriteException(exception);
            }
        }

        private void StripHost_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                ToolStripControlHost toolStripItem = sender as ToolStripControlHost;
                TrayMenuItem menuItem = toolStripItem.Tag as TrayMenuItem;
                RaiseItemKeyPress(menuItem, e);
            }
            catch (Exception exception)
            {
                NetOffice.Core.Default.Console.WriteException(exception);
            }
        }
        
        #endregion
    }
}
