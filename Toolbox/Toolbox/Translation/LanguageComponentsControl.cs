using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NetOffice.DeveloperToolbox.Controls.Tree;

namespace NetOffice.DeveloperToolbox.Translation
{
    public partial class LanguageComponentsControl : UserControl
    {
        #region API

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool EnableWindow(IntPtr hWnd, bool bEnable);

        #endregion

        #region Fields

        private ToolLanguage _selectedLanguage;
        private ToolLanguageForm _owner;
        private Control _lastHightlight;
        private LocalizableCompoment _currentComponent;
        private NotifyPropertyChanged _currentItem;

        #endregion

        #region Ctor

        public LanguageComponentsControl()
        {
            InitializeComponent();
            textBoxString.Dock = DockStyle.Fill;
            textBoxWideString.Dock = DockStyle.Fill;
            textBoxRichString.Dock = DockStyle.Fill;
        }

        #endregion

        #region Events

        public event EventHandler SelectionChanged;

        private void RaiseSelectionChanged()
        {
            if (null != SelectionChanged)
                SelectionChanged(null, EventArgs.Empty);
        }

        #endregion

        #region Properties

        internal string SelectedNodeText
        {
            get 
            {
                if (treeGridView1.SelectedRows.Count > 0)
                {
                    TreeGridNode node = treeGridView1.SelectedRows[0] as TreeGridNode;
                    return node.Cells[0].Value as string;
                }
                else
                    return null;
            }
        }

        internal ToolLanguage SelectedLanguage
        {
            get
            {
                return _selectedLanguage;
            }
            set
            {
                _selectedLanguage = value;
                ShowComponents();
            }
        }

        #endregion

        #region Methods

        internal void HandleKeyDown()
        {
            if (treeGridView1.SelectedRows.Count == 0 || treeGridView1.Focused)
                return;

            TreeGridNode selectedNode = treeGridView1.SelectedRows[0] as TreeGridNode;
            int currentRowIndex = selectedNode.Index;
            if (null != selectedNode.Parent)
            {
                TreeGridNode parentNode = selectedNode.Parent as TreeGridNode;
                if (currentRowIndex + 1 < parentNode.Nodes.Count)
                    parentNode.Nodes[currentRowIndex + 1].Cells[0].Selected = true;
            }
            else
            {
                if (currentRowIndex + 1 < treeGridView1.Rows.Count)
                    treeGridView1.Rows[currentRowIndex + 1].Selected = true;
            }

            if (textBoxString.Visible)
                textBoxString.Focus();
            else if (textBoxWideString.Visible)
                textBoxWideString.Focus();
            else if (textBoxRichString.Visible)
                textBoxRichString.Focus();
        }

        internal void HandleKeyUp()
        {
            if (treeGridView1.SelectedRows.Count == 0 || treeGridView1.Focused)
                return;

            TreeGridNode selectedNode = treeGridView1.SelectedRows[0] as TreeGridNode;
            int currentRowIndex = selectedNode.Index;
            if (null != selectedNode.Parent)
            {
                TreeGridNode parentNode = selectedNode.Parent as TreeGridNode;
                if (currentRowIndex > 0)
                    parentNode.Nodes[currentRowIndex -1].Cells[0].Selected = true;
            }
            else
            {
                if (currentRowIndex - 1 > 0)
                    treeGridView1.Rows[currentRowIndex -1].Selected = true;
            }

            if (textBoxString.Visible)
                textBoxString.Focus();
            else if (textBoxWideString.Visible)
                textBoxWideString.Focus();
            else if (textBoxRichString.Visible)
                textBoxRichString.Focus();

        }

        private void EnableInput(Control control, bool enabled)
        {
            if(control.IsDisposed || control.IsHandleCreated == false)
                return;
            IntPtr controlPtr = (IntPtr)control.Handle.ToInt32();
            EnableWindow(controlPtr, enabled);
        }

        private ToolLanguageForm FindOwner(Control ctrl)
        {
            if (null == ctrl)
                return null;

            if (ctrl.Parent is ToolLanguageForm)
                return ctrl.Parent as ToolLanguageForm;
            else
            {
                ToolLanguageForm p = FindOwner(ctrl.Parent);
                return p;
            }
        }

        private void ShowComponents()
        {
            Font boldFont = new Font(treeGridView1.DefaultCellStyle.Font, FontStyle.Bold);

            treeGridView1.Nodes.Clear();
            if (null != _selectedLanguage)
            {
                foreach (var item in _selectedLanguage.Components)
                {
                    EnableInput(item.Design, false);
                    TreeGridNode node = treeGridView1.Nodes.Add(item.Value);
                    node.DefaultCellStyle.Font = boldFont;
                    node.Tag = item;
                    node.ImageIndex = 0;
                    foreach (var subItem in item.ControlRessources)
                    {
                        string text = subItem.Value;
                        if(text.Equals("this"))
                            text = "[Title]";
                        TreeGridNode subNode = node.Nodes.Add(text);
                        subNode.Tag = subItem;
                        subNode.ImageIndex = 1;                        
                    }
                }
            }
            if (treeGridView1.SelectedRows.Count > 0)
                treeGridView1_SelectionChanged(treeGridView1, EventArgs.Empty);
        }

        private void Clear()
        {
            while (tabPage1.Controls.Count > 0)
                tabPage1.Controls.Remove(tabPage1.Controls[0]);
            textBoxString.Visible = false;
            textBoxWideString.Visible = false;
            textBoxRichString.Visible = false;
        }

        private void ShowStringEditor(LocalizableCompoment component, NotifyPropertyChanged item)
        {
            if (null != component && null != item)
            {
                if (item is LocalizableRTFString)
                {
                    textBoxRichString.RichText = item.Value2; 
                    textBoxRichString.Enabled = true;
                    textBoxRichString.Visible = true;
                }
                else if (item is LocalizableWideString)
                {
                    textBoxWideString.Text = item.Value2;
                    textBoxWideString.Enabled = true;
                    textBoxWideString.Visible = true;
                }
                else 
                {
                    textBoxString.Text = item.Value2;
                    textBoxString.Enabled = true;
                    textBoxString.Visible = true;
                }

                _currentComponent = component;
                _currentItem = item;
            }
            else
            {
                _currentComponent = null;
                _currentItem = null;
                textBoxString.Enabled = false;
                textBoxWideString.Enabled = false;
                textBoxRichString.Enabled = false;
                textBoxString.Visible = false;
                textBoxWideString.Visible = false;
                textBoxRichString.Visible = false;
                textBoxString.Text = String.Empty;
                textBoxWideString.Text = String.Empty;
                textBoxRichString.RichText = String.Empty;
            }
        }

        internal void DisableComponents()
        {
            foreach (var item in treeGridView1.Nodes)
            {
                LocalizableCompoment comp = item.Tag as LocalizableCompoment;
                EnableInput(comp.Design, false);
            }
        }

        #endregion

        #region Trigger

        private void treeGridView1_SelectionChanged(object sender, EventArgs e)
        {
            RaiseSelectionChanged();
            Clear();
            if(null == _owner)
                _owner = FindOwner(this);
            _owner.StopHighLightControl2();
            _lastHightlight = null;
           
            if (treeGridView1.SelectedCells.Count > 0)
            {
                int index = treeGridView1.SelectedCells[0].RowIndex;
                TreeGridNode node = treeGridView1.Rows[index] as TreeGridNode;
                
                LocalizableCompoment component;
                component = node.Tag as LocalizableCompoment;
                if (null == component)
                {
                    if (null != (node.Parent.Tag as LocalizableCompoment))
                    {
                        string ctrlName = (node.Tag as NotifyPropertyChanged).Value;
                        if (!ctrlName.Equals("this", StringComparison.InvariantCulture))
                        { 
                            Control ctrl = Translator.TryGetControl((node.Parent.Tag as LocalizableCompoment).Design, (node.Tag as NotifyPropertyChanged).Value);
                            if (null != ctrl && tabControl1.SelectedIndex == 0)
                            {
                                _owner.StartHighLightControl2(ctrl);
                                _lastHightlight = ctrl;
                            }
                            else
                            {
                                _owner.StopHighLightControl2();
                                _lastHightlight = null;
                            }
                        }
                    }

                    component = node.Parent.Tag as LocalizableCompoment;
                    ShowStringEditor(component, node.Tag as NotifyPropertyChanged);
                }
                else
                    ShowStringEditor(null, node.Tag as NotifyPropertyChanged);

                if (null != component)
                {
                    ILocalizationDesign design = component.Design as ILocalizationDesign;
                    if (null != design)
                    {
                        design.Localize(component.ControlRessources);
                    }
                    
                    tabPage1.Controls.Add(component.Design);
                    component.Design.Dock = DockStyle.Fill;
                    component.Design.Visible = true;
                    EnableInput(component.Design, false);
                }
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 1)
                _owner.StopHighLightControl2();
            else
                _owner.StartHighLightControl2(_lastHightlight);
        }

        private void treeGridView1_DoubleClick(object sender, EventArgs e)
        {
            if (treeGridView1.SelectedCells.Count > 0)
            {
                int index = treeGridView1.SelectedCells[0].RowIndex;
                TreeGridNode node = treeGridView1.Rows[index] as TreeGridNode;
                if (node.Nodes.Count > 0 && false == node.IsExpanded)
                    node.Expand();
                else
                { 
                    if (node.Nodes.Count > 0 && true == node.IsExpanded)
                        node.Collapse();
                }
            }
        }

        private void textBoxString_TextChanged(object sender, EventArgs e)
        {
            if (!textBoxString.Enabled || !textBoxString.Visible || null == _currentComponent || null == _currentItem)
                return;
            (_currentComponent.Design as ILocalizationDesign).Localize(_currentItem.Value, textBoxString.Text);
            _currentItem.Value2 = textBoxString.Text;
            _selectedLanguage.IsDirty = true;
        }

        private void textBoxWideString_TextChanged(object sender, EventArgs e)
        {
            if (!textBoxWideString.Enabled || !textBoxWideString.Visible || null == _currentComponent || null == _currentItem)
                return;
            (_currentComponent.Design as ILocalizationDesign).Localize(_currentItem.Value, textBoxWideString.Text);
            _currentItem.Value2 = textBoxWideString.Text;
            _selectedLanguage.IsDirty = true;
        }

        private void textBoxRichString_TextChanged(object sender, EventArgs e)
        {
            if (!textBoxRichString.Enabled || !textBoxRichString.Visible || null == _currentComponent || null == _currentItem)
                return;
            (_currentComponent.Design as ILocalizationDesign).Localize(_currentItem.Value, textBoxRichString.RichText);
            _currentItem.Value2 = textBoxRichString.RichText;
            _selectedLanguage.IsDirty = true;
        }

        #endregion
    }
}
