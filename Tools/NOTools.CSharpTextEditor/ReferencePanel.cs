using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NOTools.CSharpTextEditor
{
    public partial class ReferencePanel : UserControl
    {
        #region Ctor

        public ReferencePanel()
        {
            InitializeComponent();
        }

        public ReferencePanel(CodeEditorControl parent)
        {
            ParentEditor = parent;
            InitializeComponent();
        }

        #endregion

        #region Events

        public event EventHandler OpenHideClick;

        private void RaiseOpenHideClick()
        {
            if (null != OpenHideClick)
                OpenHideClick(this, new EventArgs());
        }

        #endregion

        #region Properties

        internal CodeEditorControl ParentEditor { get; set; }

        [Category("ReferencePanel")]
        internal bool PanelOpen
        {
            get 
            {
                return _panelOpen;
            }
            set 
            {
                _panelOpen = value;
                RaiseOpenHideClick();
            }
        }
        private bool _panelOpen = true;

        /// <summary>
        /// Current shown referencs
        /// </summary>
        internal string[] References
        {
            get
            {
                List<string> list = new List<string>();
                foreach (TreeNode item in treeView1.Nodes)
                    list.Add(item.Text);
                return list.ToArray();
            }
        }

        #endregion

        #region Methods

        internal void ShowReferences()
        {
            if (null == ParentEditor)
                return;
            if (null == ParentEditor.References)
                return;

            treeView1.Nodes.Clear();
            foreach (var item in ParentEditor.References)
            {
                treeView1.Nodes.Add(item.Name);
            }
        }

        internal void PerformHide()
        {
            labelHeader.Visible = false;
            treeView1.Visible = false;
        }

        internal void PerformVisible()
        {
            labelHeader.Visible = true;
            treeView1.Visible = true;
        }

        internal bool IsValidAssembly(string fullFileName)
        {
            try
            {
                Mono.Cecil.AssemblyDefinition.ReadAssembly(fullFileName);
                return true;
            }
            catch
            {
                return false;                
            }
        }

        #endregion

        #region Trigger

        private void buttonOpenHide_Click(object sender, EventArgs e)
        {
            _panelOpen = !_panelOpen;
            RaiseOpenHideClick();
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            if (!ParentEditor.ReferencePanelSettings.AllowAddRemoveReferences)
            { 
                e.Cancel = true;
                return;
            }
            toolStripMenuItemRemove.Enabled = (null != treeView1.SelectedNode);
        }

        private void toolStripMenuItemAdd_Click(object sender, EventArgs e)
        {
            if (null == ParentEditor)
                return;
            if (null == ParentEditor.References)
                return;

            ReferencesDialog dialog = new ReferencesDialog();
            dialog.GACTitle = ParentEditor.ReferencePanelSettings.GACTitle;
            dialog.FileSystemTitle = ParentEditor.ReferencePanelSettings.FileSystemTitle;
            dialog.OkButtonTitle = ParentEditor.ReferencePanelSettings.OkButtonTitle;
            dialog.CancelButtonTitle = ParentEditor.ReferencePanelSettings.CancelButtonTitle;
            dialog.Text = ParentEditor.ReferencePanelSettings.DialogTitle;
        
            if (DialogResult.OK == dialog.ShowDialog(this)) 
            {
                foreach (var item in dialog.SelectedAssemblies)
                { 
                    if(IsValidAssembly(item.Path))
                        ParentEditor.References.Add(new AssemblyReference(item.Name, item.Path));
                }
            }             
        }

        private void toolStripMenuItemRemove_Click(object sender, EventArgs e)
        {
            if (null == ParentEditor)
                return;
            if (null == ParentEditor.References)
                return;
            if (null == treeView1.SelectedNode)
                return;
            AssemblyReference item = ParentEditor.References[treeView1.SelectedNode.Text];
            if(null != item)
                ParentEditor.References.Remove(item);
        }

        #endregion
    }
}
