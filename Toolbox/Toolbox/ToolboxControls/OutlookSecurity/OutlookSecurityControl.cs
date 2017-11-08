using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Xml;
using System.Windows.Forms;
using NetOffice.OutlookApi.Tools.Contribution.Security;

namespace NetOffice.DeveloperToolbox.ToolboxControls.OutlookSecurity
{
    /// <summary>
    /// Suspend outlook security dialog through NetOffice.OutlookSecurity
    /// </summary>
    [RessourceTable("ToolboxControls.OutlookSecurity.Strings.txt")]
    public partial class OutlookSecurityControl : UserControl, IToolboxControl
    { 
        #region Fields

        private bool _programaticChange;
        private Exception _exception;
        private SecurityDialog _dialog;
        private SecurityDialogCheckBox _targetBox;
        private SecurityDialogLeftButton _targetButton;

        #endregion

        private Automation Suppress { get; set; }

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public OutlookSecurityControl()
        {
            try
            {
                InitializeComponent();
                if (!Program.IsDesign)
                {
                    Suppress = new Automation();
                    Suppress.OnAction += Suppress_OnAction;
                    Suppress.OnError += Suppress_OnError;
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// Last Right Click Link Label
        /// </summary>
        private LinkLabel LastClickedLinkLabel { get; set; }

        #endregion

        #region IToolboxControl

        public IToolboxHost Host { get; private set; }

        public new void KeyDown(KeyEventArgs e)
        {

        }

        public string ControlName
        {
            get { return "OutlookSecurity.OutlookSecurityControl"; }
        }

        public string ControlCaption
        {
            get { return "Outlook Security"; }
        }

        public System.ComponentModel.IContainer Components
        {
            get
            {
                return components;
            }
        }

        public Image Icon
        {
            get { return Ressources.RessourceUtils.ReadIconImageFromRessource("ToolboxControls.OutlookSecurity.Icon.ico"); }
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

        public bool SupportsHelpContent
        {
            get
            {
                return true;
            }
        }

        public void InitializeControl(IToolboxHost host)
        {
            Host = host;
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
            try
            {
                System.Xml.XmlNode node = configNode.SelectSingleNode("SupressEnabled");
                if (null == node)
                {
                    node = configNode.OwnerDocument.CreateElement("SupressEnabled");
                    node.InnerText = "false";
                    configNode.AppendChild(node);
                }
                bool mode = Convert.ToBoolean(node.Value);
                checkBoxSupressEnabled.Checked = mode;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        public void SaveConfiguration(XmlNode configNode)
        {
            try
            {
                System.Xml.XmlNode node = configNode.SelectSingleNode("SupressEnabled");
                if (null == node)
                {
                    node = configNode.OwnerDocument.CreateElement("SupressEnabled");
                    node.InnerText = BoolToString(checkBoxSupressEnabled.Checked);
                    configNode.AppendChild(node);
                }
                else
                  node.InnerText = BoolToString(checkBoxSupressEnabled.Checked);
               
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        public Stream GetHelpText()
        {
                return Ressources.RessourceUtils.ReadStream("ToolboxControls.OutlookSecurity.Info1033.rtf");
        }

        public void Release()
        {
            if (null != Suppress)
            {
                Suppress.Dispose();
                Suppress = null;
            }
        }

        #endregion

        #region Methods

        private static string BoolToString(bool b)
        {
            // not sure bool.ToString() returns something else on chinese/arabic systems
            return b ? "true" : "false";
        }

        #endregion

        #region UI Trigger

        private void LinkContextMenu_ItemClicked(object sender, ToolStripItemClickedEventArgs args)
        {
            if (null != LastClickedLinkLabel)
            {
                Clipboard.SetText(LastClickedLinkLabel.Tag as string);
            }
        }

        private void linkLabel_Clicked(object sender, EventArgs args)
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
                    LastClickedLinkLabel = sender as LinkLabel;
                    LinkContextMenu.Show(sender as Control, 0, 0);
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void checkBoxSupressEnabeld_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (_programaticChange)
                    return;               
                Suppress.Enabled = checkBoxSupressEnabled.Checked;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        #endregion

        #region Supress Trigger

        private void Supress_OnError()
        {
            try
            {
                _programaticChange = true;
                checkBoxSupressEnabled.Checked = false;
                _programaticChange = false;
                labelMessages.Text = "Error:" + _exception.Message + labelMessages.Text + Environment.NewLine;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            } 
        }

        
        private void Supress_OnAction()
        {
            try
            {
                labelMessages.Text = "Dialog:" + _dialog.Caption + " CheckBox:" + _targetBox.Caption + " Button:" + _targetButton.Caption + Environment.NewLine;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            } 
       }

        private void Suppress_OnError(Exception exception)
        {
            _exception = exception;
            if (this.InvokeRequired)
                this.Invoke(new MethodInvoker(Supress_OnError));
            else
                Supress_OnError();
        }

        private void Suppress_OnAction(SecurityDialog dialog, SecurityDialogCheckBox targetBox, SecurityDialogLeftButton targetButton)
        {
            try
            {
                _dialog = dialog;
                _targetBox = targetBox;
                _targetButton = targetButton;
                if (this.InvokeRequired)
                    this.Invoke(new MethodInvoker(Supress_OnAction));
                else
                    Supress_OnAction();
            }
            catch (Exception exception)
            {
                // avoid default error dialog because we may not in UI thread
                MessageBox.Show(this, exception.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion
    }
}