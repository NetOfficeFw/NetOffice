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
        public ReferencePanel()
        {
            InitializeComponent();
        }

        #region Events

        public event EventHandler OpenHideClick;

        private void RaiseOpenHideClick()
        {
            if (null != OpenHideClick)
                OpenHideClick(this, new EventArgs());
        }

        #endregion

        #region Properties

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
            }
        }
        private bool _panelOpen = true;

        #endregion

        #region Methods

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

        #endregion

        #region Trigger

        private void buttonOpenHide_Click(object sender, EventArgs e)
        {
            _panelOpen = !_panelOpen;
            RaiseOpenHideClick();
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            toolStripMenuItemRemove.Enabled = (null != treeView1.SelectedNode);
        }

        private void toolStripMenuItemAdd_Click(object sender, EventArgs e)
        {
            ReferencesDialog dialog = new ReferencesDialog();
            dialog.ShowDialog(this);
        }

        private void toolStripMenuItemRemove_Click(object sender, EventArgs e)
        {

        }

        #endregion
    }
}
