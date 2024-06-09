using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace AddinRemovalTool
{
    public partial class Form1 : Form
    {
        #region Ctor

        public Form1()
        {
            InitializeComponent();
            Searcher = new AddinSearcher();
            Searcher.Action += new ActionEventHandler(Searcher_Action);
            buttonRefresh_Click(buttonRefresh, new EventArgs());
        }

        #endregion

        #region Properties

        internal AddinSearcher Searcher { get; private set; }

        #endregion

        #region Methods

        private void UpdateRemoveButtonState()
        {
            if (listOverView.Items.Count > 0)
            {
                foreach (ListViewItem item in listOverView.Items)
                {
                    if (item.Checked)
                    {
                        buttonRemove.Enabled = true;
                        return;
                    }
                }
            }
            buttonRemove.Enabled = false;
        }

        #endregion

        #region Trigger

        private void Searcher_Action(ActionType type, string action)
        {
            switch (type)
            {
                case ActionType.Error:
                    MessageBox.Show(action, "Removal Tool", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                default:
                    break;
            }
        }

        private void listOverView_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            UpdateRemoveButtonState();
        }

        private void buttonRefresh_Click(object sender, EventArgs e)
        {
            Searcher.Refresh();
            listOverView.Items.Clear();
            foreach (AddinEntry item in Searcher)
            {
                ListViewItem newItem = listOverView.Items.Add(item.Name);
                newItem.SubItems.Add(item.Product);
                newItem.SubItems.Add(item.Description);
                newItem.SubItems.Add(item.ProgID);
                newItem.Checked = true;
                newItem.Tag = item;
            }
            labelAddinCount.Text = string.Format("{0} registered NetOffice Sample Addins found.", Searcher.Count);
            UpdateRemoveButtonState();
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void buttonRemove_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listOverView.Items)
            {
                AddinEntry entry = item.Tag as AddinEntry;
                if (entry.Delete())
                    item.Remove();                
            }
            buttonRefresh_Click(buttonRefresh, EventArgs.Empty);
        }

        #endregion

    }
}
