using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AddinRemovalTool
{
    public partial class Form1 : Form
    {
        AddinSearcher Searcher { get; set; }

        public Form1()
        {
            InitializeComponent();
            Searcher = new AddinSearcher();
            Searcher.Action += new ActionEventHandler(Searcher_Action);
            buttonRefresh_Click(buttonRefresh, new EventArgs());
        }

        private void buttonRefresh_Click(object sender, EventArgs e)
        {
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
            Refresh();
        }


        void Searcher_Action(ActionType type, string action)
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
    }
}
