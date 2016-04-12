using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Xml;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.OfficeUI
{
    /// <summary>
    /// Allows to analyze the office user interface object model
    /// </summary>
    [RessourceTable("ToolboxControls.OfficeUI.Strings.txt")]
    public partial class OfficeUIControl : UserControl, IToolboxControl
    {
        #region Fields
      
        private ApplicationWrapper _officeApplication;
        private  WaitControl _waitControl;
        private bool _wait;

        #endregion

        #region Construction

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public OfficeUIControl()
        {
            InitializeComponent();
            NetOffice.DebugConsole.Default.Mode = DebugConsoleMode.Console;
            NetOffice.Settings.Default.UseExceptionMessage = ExceptionMessageHandling.CopyAllInnerExceptionMessagesToTopLevelException;
            _waitControl = new WaitControl(1033);
            _waitControl.Visible = false;
            this.Controls.Add(_waitControl);
        }

        #endregion

        #region IToolboxControl

        public IToolboxHost Host { get; private set; }

        public void InitializeControl(IToolboxHost host)
        {
            Host = host;
        }

        public new void KeyDown(KeyEventArgs e)
        {

        }

        public string ControlName
        {
            get { return "OfficeUI.OfficeUIControl"; }
        }

        public string ControlCaption
        {
            get { return "Office UI"; }
        }

        public Image Icon
        {
            get { return Ressources.RessourceUtils.ReadImageFromRessource("ToolboxControls.OfficeUI.Icon.png"); }
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

        public void Activate(bool firstTime)
        {

        }

        public void Deactivated()
        {

        }

        public void LoadComplete()
        {
            
        }

        public void LoadConfiguration(XmlNode configNode)
        {
            
        }

        public void SaveConfiguration(XmlNode configNode)
        {
           
        }

        public void SetLanguage(int id)
        {
            _waitControl.CurrentLanguageID = id;
        }

        public Stream GetHelpText(int lcid)
        {
            Translation.ToolLanguage language = Host.Languages[lcid, false];
            if (null != language)
            {
                string content = language.Components["Office UI - Help"].ControlRessources["richTextBoxHelpContent"].Value2;
                return Ressources.RessourceUtils.CreateStreamFromString(content);
            }
            else
                return Ressources.RessourceUtils.ReadStream("ToolboxControls.OfficeUI.Info" + lcid.ToString() + ".rtf");
        }

        public void Release()
        {
            DisposeCurrentOpenOfficeApplication();
        }

        public IContainer Components
        {
            get
            {
                return components;
            }
        }

        #endregion

        #region ILocalizationDesign

        public void EnableDesignView(int lcid, string parentComponentName)
        {

        }

        public void Localize(Translation.ItemCollection strings)
        {
            Translation.Translator.TranslateControls(this, strings);
        }

        public void Localize(string name, string text)
        {
            Translation.Translator.TranslateControl(this, name, text);
        }

        public string GetCurrentText(string name)
        {
            return Translation.Translator.TryGetControlText(this, name);
        }

        public string NameLocalization
        {
            get
            {
                return null;
            }
        }

        public IEnumerable<ILocalizationChildInfo> Childs
        {
            get
            {
                return new ILocalizationChildInfo[] { new LocalizationDefaultChildInfo("Select App", typeof(SelectOfficeAppControl)), new LocalizationDefaultChildInfo("Help", typeof(Controls.InfoLayer.InfoControl)) };
            }
        }

        #endregion

        #region Methods
        
        private void DisposeCurrentOpenOfficeApplication()
        {
            if (null != _officeApplication)
            {
                _officeApplication.Quit();
                _officeApplication.Dispose();
                _officeApplication = null;
            }
            buttonCloseOfficeApp.Enabled = false;
            propertyGridItems.SelectedObject = null;
            treeViewOfficeUI.Nodes.Clear();
        }

        private void Run(string officeAppName)
        {
            try
            {
                _wait = true;
                ShowWaitPanel(true);

                treeViewOfficeUI.Nodes.Clear();

                _officeApplication = new ApplicationWrapper(officeAppName);
                buttonCloseOfficeApp.Enabled = true;
                foreach (OfficeApi.CommandBar commandBar in _officeApplication.CommandBars)
                {
                    string itemName = commandBar.Name;
                    TreeNode node = treeViewOfficeUI.Nodes.Add(itemName);
                    _waitControl.ReportProgress(itemName);
                    node.ImageIndex = 0;
                    node.Tag = commandBar;
                    node.Nodes.Add("#stub");
                }
            }
            catch (Exception exception)
            {
                throw exception;
            }
            finally
            {
                HideWaitPanel();
                _wait = false;
            }
        }

        private void ShowWaitPanel(bool bigMode)
        {
            if (bigMode)
            {
                Cursor = Cursors.WaitCursor;
                _waitControl.Dock = DockStyle.Fill;
                _waitControl.BringToFront();
                _waitControl.Show();
                _waitControl.Refresh();
            }
            else
            {
                Cursor = Cursors.WaitCursor;
                _waitControl.Dock = DockStyle.None;
                _waitControl.Location = splitContainer1.Panel2.Location;
                _waitControl.Size = splitContainer1.Panel2.Size;
                _waitControl.Top += splitContainer1.Top;
                _waitControl.BringToFront();
                _waitControl.Show();
                _waitControl.Refresh();
            }
        }

        private void HideWaitPanel()
        {
            Cursor = Cursors.Default;
            _waitControl.Hide();
        }

        #endregion

        #region Trigger

        private void buttonStartApplication_Click(object sender, EventArgs e)
        {
            try
            {
                DisposeCurrentOpenOfficeApplication();
                SelectOfficeAppControl selectBox = new SelectOfficeAppControl(Host.CurrentLanguageID, new SelectOfficeEventHandler(Run));
                this.Controls.Add(selectBox);
                selectBox.Dock = DockStyle.Fill;
                selectBox.BringToFront();
                selectBox.Show();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }

        private void treeViewOfficeUI_AfterSelect(object sender, TreeViewEventArgs e)
        {
            try
            {
                if (!checkBoxScanForProperties.Checked)
                {
                    propertyGridItems.SelectedObject = null;
                    return;
                }

                if (e.Node.Tag is OfficeApi.CommandBar)
                {
                    if (!_wait)
                        ShowWaitPanel(false);
                    OfficeApi.CommandBar commandBar = e.Node.Tag as OfficeApi.CommandBar;
                    propertyGridItems.SelectedObject = commandBar;
                    if (!_wait)
                        HideWaitPanel();
                }
                else if (e.Node.Tag is OfficeApi.CommandBarControl)
                {
                    if (!_wait)
                        ShowWaitPanel(false);
                    OfficeApi.CommandBarControl commandBarControl = e.Node.Tag as OfficeApi.CommandBarControl;
                    propertyGridItems.SelectedObject = commandBarControl;
                    if (!_wait)
                        HideWaitPanel();
                }
                else
                    propertyGridItems.SelectedObject = null;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
            finally
            {
                if (!_wait)
                    HideWaitPanel();
            }
        }

        private void treeViewOfficeUI_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            try
            {
                if ((e.Node.Nodes.Count == 1) && (e.Node.Nodes[0].Text == "#stub"))
                {
                    ShowWaitPanel(false);

                    e.Node.Nodes.Clear();

                    if (e.Node.Tag is OfficeApi.CommandBar)
                    {
                        OfficeApi.CommandBar commandBar = e.Node.Tag as OfficeApi.CommandBar;
                        foreach (OfficeApi.CommandBarControl control in commandBar.Controls)
                        {
                            TreeNode subNode = e.Node.Nodes.Add(control.Caption);
                            subNode.ImageIndex = 1;
                            subNode.Tag = control;
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
            finally
            {
                HideWaitPanel();
            }
        }

        private void toolStripReset_Click(object sender, EventArgs e)
        {
            try
            {
                if (treeViewOfficeUI.SelectedNodes.Count == 0)
                    return;

                foreach (TreeNode node in treeViewOfficeUI.SelectedNodes)
                {
                    if (node.Tag is OfficeApi.CommandBar)
                    {
                        OfficeApi.CommandBar commandBar = node.Tag as OfficeApi.CommandBar;
                        commandBar.Reset();
                    }
                    else if (node.Tag is OfficeApi.CommandBarControl)
                    {
                        OfficeApi.CommandBarControl control = node.Tag as OfficeApi.CommandBarControl;
                        control.Reset();
                    }
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }

        private void toolStripDelete_Click(object sender, EventArgs e)
        {
            try
            {
                if (treeViewOfficeUI.SelectedNodes.Count == 0)
                    return;

                List<TreeNode> listDelete = new List<TreeNode>();
                foreach (TreeNode node in treeViewOfficeUI.SelectedNodes)
                {
                    if (node.Tag is OfficeApi.CommandBar)
                    {
                        OfficeApi.CommandBar commandBar = node.Tag as OfficeApi.CommandBar;
                        commandBar.Delete();
                        listDelete.Add(node);
                    }
                    else if (node.Tag is OfficeApi.CommandBarControl)
                    {
                        OfficeApi.CommandBarControl control = node.Tag as OfficeApi.CommandBarControl;
                        control.Delete();
                        listDelete.Add(node);
                    }
                }

                foreach (TreeNode node in listDelete)
                    node.Remove();

            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }

        private void buttonCloseOfficeApp_Click(object sender, EventArgs e)
        {
            try
            {
                DisposeCurrentOpenOfficeApplication();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception, ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }

        #endregion
    }
}
