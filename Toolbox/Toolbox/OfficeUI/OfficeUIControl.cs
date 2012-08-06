using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Xml;
using System.Text;
using System.Diagnostics;
using System.Windows.Forms;
using NetOffice;

namespace NetOffice.DeveloperToolbox
{
    public partial class OfficeUIControl : UserControl, IToolboxControl
    {
        #region Fields

        private int _currentLanguageID = 1033;
        ApplicationWrapper _officeApplication;
        WaitControl _waitControl;
        private bool _wait;

        #endregion

        #region Construction

        public OfficeUIControl()
        {
            InitializeComponent();
            NetOffice.DebugConsole.Mode = ConsoleMode.Console;
            NetOffice.Settings.UseExceptionMessage = ExceptionMessageHandling.CopyAllInnerExceptionMessagesToTopLevelException;
            _waitControl = new WaitControl(_currentLanguageID);
            _waitControl.Visible = false;
            this.Controls.Add(_waitControl);
        }

        #endregion

        #region IToolboxControl Member

        public string ControlName
        {
            get { return "OfficeUI"; }
        }

        public string ControlCaption
        {
            get { return "Office UI"; }
        }

        public Image Icon
        {
            get
            {
                return ReadImageFromRessource("Icon.png");
            }
        }

        public void Activate()
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
            _currentLanguageID = id;
            _waitControl.CurrentLanguageID = id;
            Translator.TranslateControls(this, "OfficeUI.MessageTable.txt", _currentLanguageID);
        }

        public new void Dispose()
        {
            DisposeCurrentOpenOfficeApplication();
            base.Dispose();
        }

        public IContainer Components
        {
            get
            {
                return components;
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
                SelectOfficeAppControl selectBox = new SelectOfficeAppControl(_currentLanguageID, new SelectOfficeEventHandler(Run));
                this.Controls.Add(selectBox);
                selectBox.Dock = DockStyle.Fill;
                selectBox.BringToFront();
                selectBox.Show();
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
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
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
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
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
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
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
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
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void buttonCloseOfficeApp_Click(object sender, EventArgs e)
        {
            DisposeCurrentOpenOfficeApplication();
        }

        private void buttonInfo_Click(object sender, EventArgs e)
        {
            try
            {
                InfoControl infoBox = new InfoControl("OfficeUI.Info" + _currentLanguageID.ToString() + ".rtf", true);
                this.Controls.Add(infoBox);
                infoBox.BringToFront();
                infoBox.Show();
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        #endregion
         
        #region Static Helper
         
        private static Image ReadImageFromRessource(string ressourcePath)
        {
            System.IO.Stream ressourceStream = null;
            string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            ressourcePath = assemblyName + ".OfficeUI." + ressourcePath;
            ressourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(ressourcePath);
            if (ressourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));
            Bitmap newIcon = new Bitmap(ressourceStream);
            return newIcon;
        }

        #endregion

    
    }
}
