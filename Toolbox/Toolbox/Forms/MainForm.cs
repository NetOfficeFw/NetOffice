using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using System.Xml;
using System.Windows.Forms;
using NetOffice.DeveloperToolbox.ToolboxControls;
using NetOffice.DeveloperToolbox.Utils.Native;

namespace NetOffice.DeveloperToolbox.Forms
{
    /// <summary>
    /// Main form in the application
    /// </summary>
    public partial class MainForm : Form, IToolboxHost
    {
        #region Fields

        /// <summary>
        /// the currenty loaded toolbox controls
        /// </summary>
        private List<IToolboxControl> _toolboxControls;

        /// <summary>
        /// toolbox controls with first show passed away (no dictionary for some reasons)
        /// </summary>
        private List<IToolboxControl> _toolBoxControlsFirstShowPassed;

        /// <summary>
        /// store last selection to call IToolboxControl.Deactivated() in SelectedIndexChanged
        /// </summary>
        private IToolboxControl _lastSelectedcontrol;

        /// <summary>
        /// An error occured in WndProc
        /// </summary>
        private bool _wndProcErrorOccured;

        #endregion

        #region Construction

        /// <summary>
        /// Stub Ctor
        /// </summary>
        public MainForm(): this(new string[0])
        {
            InitializeComponent();
            Singleton = this;
        } 

        /// <summary>
        /// Runtime Ctor
        /// </summary>
        /// <param name="args">commandline argument array</param>
        public MainForm(string[] args)
        {
            InitializeComponent();
            Singleton = this;
            CommandLineArgs = args;
            LoadRuntimeControls();
            LoadConfiguration();
        }

        #endregion
  
        #region Properties

        /// <summary>
        /// Commandline arguments
        /// </summary>
        private string[] CommandLineArgs { get; set; }

        /// <summary>
        /// Shared singleton instance in the AppDomain
        /// </summary>
        internal static MainForm Singleton { get; private set; }

        #endregion

        #region IToolboxHost

        public event EventHandler Minimized;

        private void RaiseMinimized()
        {
            if (null != Minimized)
                Minimized(this, EventArgs.Empty);
        }

        public string Caption
        {
            get 
            {
                return System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            }
        }
    
        public IToolboxControl[] ToolboxControls
        {
            get 
            {
                return _toolboxControls.ToArray();
            }
        }

        public void ShowMainWindow()
        {
            WindowState = FormWindowState.Normal;
            Activate();
            ShowInTaskbar = true;
        }

        public void MinimizeMainWindow(bool showInTaskbar)
        {
            WindowState = FormWindowState.Minimized;
            ShowInTaskbar = showInTaskbar;
        }
        
        public void SwitchTo(string controlName)
        {
            foreach (TabPage item in tabControlMain.TabPages)
            {
                IToolboxControl ctrl = item.Tag as IToolboxControl;
                if (null != ctrl && ctrl.ControlName == controlName)
                {
                    tabControlMain.SelectedTab = item;
                    return;
                }
            }
        }

        #endregion

        #region Methods


        private void LoadRuntimeControls()
        {
            try
            {
                //_isCurrentlyLoading = true;
                tabControlMain.TabPages.Clear();
                _toolboxControls = new List<IToolboxControl>();
                _toolBoxControlsFirstShowPassed = new List<IToolboxControl>();
                LoadRuntimeControl(typeof(ToolboxControls.Welcome.WelcomeControl));
                LoadRuntimeControl(typeof(ToolboxControls.OfficeCompatibility.OfficeCompatibilityControl));
                LoadRuntimeControl(typeof(ToolboxControls.ApplicationObserver.ApplicationObserverControl));
                LoadRuntimeControl(typeof(ToolboxControls.RegistryEditor.RegistryEditorControl));
                LoadRuntimeControl(typeof(ToolboxControls.OfficeUI.OfficeUIControl));
                LoadRuntimeControl(typeof(ToolboxControls.OutlookSecurity.OutlookSecurityControl));
                LoadRuntimeControl(typeof(ToolboxControls.ProjectWizard.ProjectWizardControl));
                LoadRuntimeControl(typeof(ToolboxControls.ProxyView.ProxyControl));
                LoadRuntimeControl(typeof(ToolboxControls.About.AboutControl));
            }
            catch
            {
                throw;
            }
            finally
            {
                //_isCurrentlyLoading = false;
            }
        }

        private void LoadRuntimeControl(Type type)
        {
            IToolboxControl ctrl = Activator.CreateInstance(type) as IToolboxControl;
            ControlContainer hostContainer = new ControlContainer(ctrl);
            tabControlMain.TabPages.Add(hostContainer.ControlCaption);
            TabPage page = tabControlMain.TabPages[tabControlMain.TabPages.Count - 1];
            page.BackColor = Color.FromArgb(0, 201, 227, 243);
            page.Margin = new System.Windows.Forms.Padding(0);
            page.Padding = new System.Windows.Forms.Padding(0);
            page.Controls.Add(hostContainer as Control);
            page.Tag = hostContainer;
            imageListTabMain.Images.Add(hostContainer.Icon);
            page.ImageIndex = imageListTabMain.Images.Count - 1;
            (hostContainer as Control).Dock = DockStyle.Fill;
            _toolboxControls.Add(hostContainer);
            hostContainer.InitializeControl(this);
        }

        private XmlDocument CreateDefaultConfiguration()
        {
            XmlDocument document = new XmlDocument();
            XmlNode root = document.CreateNode(XmlNodeType.Element, "NODeveloperToolbox.Settings", null) as XmlNode;
            XmlAttribute versionAttribute = document.CreateAttribute("Version");
            versionAttribute.Value = AssemblyInfo.AssemblyVersion;
            root.Attributes.Append(versionAttribute);

            document.AppendChild(root);
            foreach (var item in _toolboxControls)
            {
                XmlNode configNode = document.CreateNode(XmlNodeType.Element, item.ControlName, null);
                root.AppendChild(configNode);
            }
            return document;
        }

        private void LoadConfiguration()
        {           
            XmlDocument document = null;
            string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "NODeveloperToolbox.Settings.xml");
            if (File.Exists(filePath))
            {
                document = new XmlDocument();
                document.Load(filePath);
                XmlAttribute versionAttribute = document.FirstChild.Attributes["Version"];
                if (null != versionAttribute && document.FirstChild.LocalName == "NODeveloperToolbox.Settings")
                { 
                    string configVersion = versionAttribute.Value;
                    if (!configVersion.Equals(AssemblyInfo.AssemblyVersion, StringComparison.InvariantCultureIgnoreCase))
                        document = CreateDefaultConfiguration();
                }
                else
                    document = CreateDefaultConfiguration();
            }
            else
                document = CreateDefaultConfiguration();

            foreach (var item in _toolboxControls)
                item.LoadConfiguration(document.SelectSingleNode("NODeveloperToolbox.Settings/" + item.ControlName));
        }

        private void SaveConfiguration()
        {
            XmlDocument document = new XmlDocument();
            string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "NODeveloperToolbox.Settings.xml");
            XmlNode root = document.CreateNode(XmlNodeType.Element, "NODeveloperToolbox.Settings", null);
            XmlAttribute versionAttribute = document.CreateAttribute("Version");
            versionAttribute.Value = AssemblyInfo.AssemblyVersion;
            root.Attributes.Append(versionAttribute);
            document.AppendChild(root);
            foreach (var item in _toolboxControls)
            {
                XmlNode configNode = document.CreateNode(XmlNodeType.Element, item.ControlName, null);
                item.SaveConfiguration(configNode);
                root.AppendChild(configNode);
            }

            document.Save(filePath);
        }

        #endregion

        #region Overrides

        /// <summary>
        /// We wait for ouer custom message to bring up the window in front 
        /// </summary>
        /// <param name="m">window message kind and args </param>
        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
           
            if (_wndProcErrorOccured)
                return;
            
            try
            {
                if (m.Msg == Win32.WM_SHOWTOOLBOX)
                {
                    if (WindowState == FormWindowState.Minimized)
                        WindowState = FormWindowState.Normal;
                    Win32.SetForegroundWindow(Handle);
                    Show();
                    BringToFront();
                    bool top = TopMost;
                    TopMost = true;
                    TopMost = top;
                }
            }
            catch
            {
                // dont show errors here. may its called multiple times in a milliseconds
                // we set a flag and dont try to run the code again because its only a minor
                _wndProcErrorOccured = true;
            }
        }

        #endregion

        #region UI Trigger

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                bool handled = false;
                if (e.Modifiers == Keys.Alt)
                {
                    // use Alt + number key as tab change command
                    switch (e.KeyCode)
                    {
                        case Keys.D1:
                            SwitchTo(_toolboxControls[0].ControlName);
                            handled = true;
                            break;
                        case Keys.D2:
                            SwitchTo(_toolboxControls[1].ControlName);
                            handled = true;
                            break;
                        case Keys.D3:
                            SwitchTo(_toolboxControls[2].ControlName);
                            handled = true;
                            break;
                        case Keys.D4:
                            SwitchTo(_toolboxControls[3].ControlName);
                            handled = true;
                            break;
                        case Keys.D5:
                            SwitchTo(_toolboxControls[4].ControlName);
                            handled = true;
                            break;
                        case Keys.D6:
                            SwitchTo(_toolboxControls[5].ControlName);
                            handled = true;
                            break;
                        case Keys.D7:
                            SwitchTo(_toolboxControls[6].ControlName);
                            handled = true;
                            break;
                        case Keys.D8:
                            SwitchTo(_toolboxControls[7].ControlName);
                            handled = true;
                            break;
                    }
                }

                // no tab change command, we shift key input to current visible toolbox control(s)
                if (!handled)
                {
                    foreach (var item in ToolboxControls)
                    {
                        Control winControl = item as Control;
                        if (winControl.Visible)
                        {
                            item.KeyDown(e);
                            return;
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        private void TabControlMain_Deselecting(object sender, TabControlCancelEventArgs e)
        {
            try
            {
                IToolboxControl lastControl = null;

                if (null != tabControlMain.SelectedTab)
                    lastControl = tabControlMain.SelectedTab.Tag as IToolboxControl;

                _lastSelectedcontrol = lastControl;
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        private void TabControlMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                TabPage currentTab = tabControlMain.TabPages[tabControlMain.SelectedIndex];
                IToolboxControl control = currentTab.Tag as IToolboxControl;
                if (null != control)
                {
                    bool firstShow = !_toolBoxControlsFirstShowPassed.Any(c => c == control);
                    control.Activate(firstShow);
                    if (!firstShow)
                        _toolBoxControlsFirstShowPassed.Add(control);
                    if (null != _lastSelectedcontrol)
                    {
                        _lastSelectedcontrol.Deactivated();
                        _lastSelectedcontrol = null;
                    }
                }
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);        
            }
        }

        private void MainForm_Resize(object sender, EventArgs e)
        {
            try
            {
                if ((FormWindowState.Minimized == this.WindowState))
                    RaiseMinimized();
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                SaveConfiguration();
                foreach (IToolboxControl item in ToolboxControls)
                {
                    item.Release();
                    item.Dispose();
                }
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                foreach (IToolboxControl item in ToolboxControls)
                    item.LoadComplete();
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }
   
        private void MainForm_Shown(object sender, EventArgs e)
        {
            try
            {
                _toolboxControls[0].Activate(true);
                _toolBoxControlsFirstShowPassed.Add(_toolboxControls[0]);
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        #endregion
    }
}
