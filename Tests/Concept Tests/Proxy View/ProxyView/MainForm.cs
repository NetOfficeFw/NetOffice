using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NetOffice;

namespace ProxyView
{
    public partial class MainForm : Form
    {
        #region Ctor

        public MainForm()
        {
            InitializeComponent();
            LoadSettings();
            RotDataSource = new RotEntryCollection();
            WndDataSource = new WindowEntryCollection();
            SetDataSource();
            RefreshDataSource();            
            ShowEntryDetails();
            UpdateFormCaption();
        }

        #endregion

        #region Properties

        private RotEntryCollection RotDataSource { get; set; }

        private WindowEntryCollection WndDataSource { get; set; }

        private IRefresh CurrentDataSource { get; set; }

        private Entry SelectedItem
        {
            get
            {
                if (EntriesGrid.SelectedCells.Count == 0)
                    return null;
                Entry selectedItem = EntriesGrid.Rows[EntriesGrid.SelectedCells[0].RowIndex].DataBoundItem as Entry;
                return selectedItem;
            }
        }

        #endregion

        #region Methods

        private void LoadSettings()
        {
            int interval = Properties.Settings.Default.RefreshInterval;
            if (interval < 1000)
                interval = 1000;
            if (interval > 90000)
                interval = 90000;
            RefreshTimer.Interval = interval;
        }

        private void RefreshDataSource()
        {
            if(null != CurrentDataSource)
                CurrentDataSource.Refresh();
        }

        private void SetDataSource()
        {
            if (ShowAccessibleInsteadToolStripMenuItem.Checked)
                CurrentDataSource = WndDataSource;
            else
                CurrentDataSource = RotDataSource;
            EntriesGrid.DataSource = CurrentDataSource;
        } 

        private void ShowEntryDetails()
        {
            if (Properties.Settings.Default.ShowDetails == false || EntriesGrid.SelectedCells.Count == 0)
            {
                EntryDetailsGrid.SelectedObject = null;
                return;
            }

            Entry selectedItem = SelectedItem;
            try
            {
                EntryDetailsGrid.SelectedObject = null != selectedItem ? selectedItem.Underlying : null;
            }
            catch (System.Runtime.InteropServices.InvalidComObjectException)
            {                
                EntryDetailsGrid.SelectedObject = null;
                ; // COM object that has been separated from its underlying RCW cannot be used
            }
            catch(Exception)
            {
                throw;
            }
        }

        private void UpdateFormCaption()
        {           
            Text = String.Format(
                "{0} - {1} [{2}]",
                AboutForm.AssemblyTitle, 
                CurrentDataSource == RotDataSource ? "Running Object Table" : "IAccessible From Desktop",
                Program.IsAdmin == true ? "Currently Admin Permissions" : "Currently No Admin Permissions");            
        }

        #endregion

        #region Trigger

        private void RefreshTimer_Tick(object sender, EventArgs e)
        {
            if (CurrentDataSource.IsCurrentlyRefresh)
                return;
            BeginInvoke(new Action(RefreshDataSource));
        }

        private void EntriesGrid_SelectionChanged(object sender, EventArgs e)
        {
            ShowEntryDetails();
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void RefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RefreshDataSource();
        }

        private void InfoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutForm.ShowForm(this);
        }

        private void AutoRefreshToolStripMenuItem_CheckedChanged(object sender, EventArgs e)
        {
            RefreshTimer.Enabled = AutoRefreshToolStripMenuItem.Checked;
        }

        private void ShowAccessibleInsteadToolStripMenuItem_CheckedChanged(object sender, EventArgs e)
        {
            SetDataSource();
            RefreshDataSource();
            ShowEntryDetails();
            UpdateFormCaption();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (null != RotDataSource)
                RotDataSource.Dispose();
            if (null != WndDataSource)
                WndDataSource.Dispose();
        }

        private void GridContextMenu_Opening(object sender, CancelEventArgs e)
        {
            Entry selectedItem = SelectedItem;
            if(null == selectedItem)
            {
                e.Cancel = true; 
                return;
            }
            GridContextMenu.Items.Clear();
            foreach (DataGridViewColumn column in EntriesGrid.Columns)
            {
                ToolStripItem stripItem = GridContextMenu.Items.Add(String.Format("Copy {0}", column.HeaderText));
                stripItem.Tag = (CurrentDataSource as ITypedList).GetItemProperties(null).Find(column.HeaderText, true);
            }
            GridContextMenu.Items.Add(new ToolStripSeparator());
            GridContextMenu.Items.Add("Copy Entire Row");
        }

        private void GridContextMenu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            Entry selectedItem = SelectedItem;
            if (null == selectedItem)
                return;

            PropertyDescriptor descriptor = e.ClickedItem.Tag as PropertyDescriptor;
            if (null != descriptor)
            {
                object value = descriptor.GetValue(selectedItem);
                if (null != value)
                    Clipboard.SetText(value.ToString());
            }
            else
            {
                string line = String.Empty;
                foreach (PropertyDescriptor item in (CurrentDataSource as ITypedList).GetItemProperties(null))
                {
                    object value = item.GetValue(selectedItem);
                    if (null != value)
                        line += String.Format("{0} ", value);
                }
                if (!String.IsNullOrWhiteSpace(line))
                    Clipboard.SetText(line);
            }
        }

        private void AdvancedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (SettingsForm.ShowForm(this) == DialogResult.OK)
            {
                RefreshDataSource();
                RefreshTimer.Interval = Properties.Settings.Default.RefreshInterval;
                ShowEntryDetails();                
            }
        }

        #endregion
    }
}
