using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.About
{
    /// <summary>
    /// Application about panel
    /// </summary>
    [RessourceTable("ToolboxControls.About.Strings.txt")]
    public partial class AboutControl : UserControl, IToolboxControl
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public AboutControl()
        {
            InitializeComponent();
            labelVersionText.Text = String.Format("Version {0}", AssemblyInfo.AssemblyVersion);
            labelCopyrightText.Text = AssemblyInfo.AssemblyCopyright;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Last Right Click Link Label
        /// </summary>
        private Control LastClickedLinkLabel { get; set; }

        #endregion

        #region IToolboxControl

        public IToolboxHost Host { get; private set; }

        public string ControlName
        {
            get { return "About.AboutControl"; }
        }

        public string ControlCaption
        {
            get { return "About"; }
        }

        public Image Icon
        {
            get { return Ressources.RessourceUtils.ReadImageFromRessource("ToolboxControls.About.info_rhombus.png"); }
        }

        public bool SupportsHelpContent
        {
            get
            {
                return false;
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
            controlForeColorAnimator1.Start(false);
        }

        public void Deactivated()
        {
            controlForeColorAnimator1.Stop();
        }

        public void LoadComplete()
        {

        }

        public void LoadConfiguration(System.Xml.XmlNode configNode)
        {

        }

        public void SaveConfiguration(System.Xml.XmlNode configNode)
        {

        }

        public Stream GetHelpText()
        {
            throw new NotImplementedException();
        }

        public new void KeyDown(KeyEventArgs e)
        {

        }

        public void Release()
        {

        }

        public IContainer Components
        {
            get { return components; }
        }

        #endregion

        #region Trigger

        private void labelNetOfficeIsFree_Click(object sender, EventArgs args)
        {
            try
            {
                MouseEventArgs mouseArgs = args as MouseEventArgs;
                if (null != mouseArgs && mouseArgs.Button == MouseButtons.Left)
                {
                    Control control = sender as Control;
                    System.Diagnostics.Process.Start(control.Tag as string);
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void LinkContextMenu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (null != LastClickedLinkLabel)
            {
                Clipboard.SetText(LastClickedLinkLabel.Tag as string);
            }
        }

        private void linkLabelMailContact_Click(object sender, EventArgs args)
        {
            try
            {
                MouseEventArgs mouseArgs = args as MouseEventArgs;
                if (null == mouseArgs)
                    return;

                if (mouseArgs.Button == MouseButtons.Left)
                {
                    LinkLabel label = sender as LinkLabel;
                    System.Diagnostics.Process.Start("mailto:" + label.Text as string);
                }
                else if (mouseArgs.Button == MouseButtons.Right)
                {
                    LastClickedLinkLabel = sender as Control;
                    LinkContextMenu.Show(sender as Control, 0, 0);
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void linkLabelCompany_Clicked(object sender, EventArgs args)
        {
            try
            {
                MouseEventArgs mouseArgs = args as MouseEventArgs;
                if (null == mouseArgs)
                    return;

                if (mouseArgs.Button == MouseButtons.Left)
                {
                    LinkLabel label = sender as LinkLabel;
                    System.Diagnostics.Process.Start(label.Tag as string);
                }
                else if (mouseArgs.Button == MouseButtons.Right)
                {
                    LastClickedLinkLabel = sender as Control;
                    LinkContextMenu.Show(sender as Control, 0, 0);
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void AboutControl_Resize(object sender, EventArgs e)
        {
            try
            {
                panelMain.Location = new Point((this.Width / 2) - (panelMain.Width / 2), (this.Height / 2) - (panelMain.Height / 2));
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        #endregion
    }
}