using System;
using System.IO;
using System.Xml;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProxyView
{
    public partial class ProxyControl : UserControl, IToolboxControl
    {
        public ProxyControl()
        {
            InitializeComponent();
        }

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

        private void DataSourceRefreshComplete(IRefresh sender)
        {

        }

        private void LoadSettings()
        {
            int interval = Settings.RefreshInterval;
            if (interval < 1000)
                interval = 1000;
            if (interval > 90000)
                interval = 90000;
            RefreshTimer.Interval = interval;
        }

        private void RefreshDataSource()
        {
            if (null != CurrentDataSource)
                CurrentDataSource.RefreshAsync(DataSourceRefreshComplete, this);
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
            if (Settings.ShowDetails == false || EntriesGrid.SelectedCells.Count == 0)
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
            catch (Exception)
            {
                ;
            }
        }

        #endregion

        #region IToolboxControl

        public IToolboxHost Host { get; private set; }

        public string ControlName
        {
            get { return "ProxyView.ProxyControl"; }
        }

        public string ControlCaption
        {
            get { return "Proxy View"; }
        }

        public Image Icon
        {
            get { return Ressources.RessourceUtils.ReadImageFromRessource("ToolboxControls.ProxyView.ProxyView.png"); }
        }

        public bool SupportsHelpContent
        {
            get
            {
                return true;
            }
        }

        public bool SupportsInfoMessage
        {
            get
            {
                return false;
            }
        }

        public ToolboxControlMessageKind InfoMessageKind
        {
            get
            {
                return ToolboxControlMessageKind.Uncategorized;
            }
        }

        public string InfoMessage
        {
            get
            {
                return String.Empty;
            }
        }

        public void InitializeControl(IToolboxHost host)
        {
            Host = host;
        }

        public void Activate(bool firstTime)
        {
            if (firstTime)
            {
                RotDataSource = new RotEntryCollection();
                WndDataSource = new WindowEntryCollection();
                SetDataSource();
                RefreshDataSource();
                ShowEntryDetails();
            }
        }

        public void Deactivated()
        {

        }

        public Stream GetHelpText()
        {
            return Ressources.RessourceUtils.ReadStream("ToolboxControls.ProxyView.Info1033.rtf");
        }

        public void LoadComplete()
        {

        }

        public void LoadConfiguration(System.Xml.XmlNode configNode)
        {
            bool value = false;
            System.Xml.XmlNode node = configNode["ShowSelectedDetails"];
            if (null != node && bool.TryParse(node.InnerText, out value))
            {
                Settings.ShowDetails = value;
            }

            node = configNode["ShowAllAccessible"];
            if (null != node && bool.TryParse(node.InnerText, out value))
            {
                Settings.ShowAllAccessible = value;
            }
        }

        public void SaveConfiguration(System.Xml.XmlNode configNode)
        {
            System.Xml.XmlNode node = configNode.OwnerDocument.CreateElement("ShowSelectedDetails");
            configNode.AppendChild(node);
            node.InnerText = Settings.ShowDetails.ToString();

            node = configNode.OwnerDocument.CreateElement("ShowAllAccessible");
            configNode.AppendChild(node);
            node.InnerText = Settings.ShowAllAccessible.ToString();
        }

        public new void KeyDown(KeyEventArgs e)
        {

        }

        public void Release()
        {

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

        private void RefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RefreshDataSource();
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
            if (null == selectedItem)
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
                RefreshTimer.Interval = Settings.RefreshInterval;
                ShowEntryDetails();
            }
        }

        #endregion
    }
}