using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Xml;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.OutlookSecurity
{
    public partial class OutlookSecurityControl : UserControl, IToolboxControl
    { 
        #region Fields

        private int _currentLanguageID = 1033;
        private bool _programaticChange;
        private Exception _exception;
        private NetOffice.OutlookSecurity.SecurityDialog _dialog;
        private NetOffice.OutlookSecurity.SecurityDialogCheckBox _targetBox;
        private NetOffice.OutlookSecurity.SecurityDialogLeftButton _targetButton;

        #endregion

        #region Construction

        public OutlookSecurityControl()
        {
            try
            {
                InitializeComponent();
                if (!DesignMode)
                {
                    NetOffice.OutlookSecurity.Supress.OnAction += new NetOffice.OutlookSecurity.Supress.SecurityDialogAction(Supress_OnAction);
                    NetOffice.OutlookSecurity.Supress.OnError += new NetOffice.OutlookSecurity.Supress.ErrorOccuredEventHandler(Supress_OnError);
                }
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        #endregion

        #region Properties

        private new bool DesignMode
        {
            get
            {
                return (System.Diagnostics.Process.GetCurrentProcess().ProcessName == "devenv");
            }
        }
       
        #endregion

        #region IToolboxControl Member

        public string ControlName
        {
            get { return "OutlookSecurity"; }
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
            get { return ReadImageFromRessource("Icon.ico"); }
        }

        public void Activate()
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
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        public void SaveConfiguration(XmlNode configNode)
        {
            try
            {
                System.Xml.XmlNode node = configNode.SelectSingleNode("SupressEnabled");
                node.InnerText = BoolToString(checkBoxSupressEnabled.Checked);
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        public void SetLanguage(int id)
        {
            _currentLanguageID = id;
            Translator.TranslateControls(this, "OutlookSecurity.MessageTable.txt", _currentLanguageID);
        }

        #endregion

        #region UI Trigger

        private void linkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                LinkLabel label = sender as LinkLabel;
                System.Diagnostics.Process.Start(label.Text);
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void checkBoxSupressEnabeld_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (_programaticChange)
                    return;
                NetOffice.OutlookSecurity.Supress.Enabled = checkBoxSupressEnabled.Checked;
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void buttonInfo_Click(object sender, EventArgs e)
        {
            try
            {
                InfoControl infoBox = new InfoControl("OutlookSecurity.Info" + _currentLanguageID.ToString() + ".rtf", true);
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

        #region Supress Trigger

        void Supress_OnError()
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
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            } 
        }
    
        void Supress_OnError(Exception exception)
        {
            _exception= exception;
            if (this.InvokeRequired)
                this.Invoke(new MethodInvoker(Supress_OnError));
            else
                Supress_OnError();            
        }

        void Supress_OnAction()
        {
            try
            {
                labelMessages.Text = "Dialog:" + _dialog.Caption + " CheckBox:" + _targetBox.Caption + " Button:" + _targetButton.Caption + Environment.NewLine;
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            } 
       }

        void Supress_OnAction(NetOffice.OutlookSecurity.SecurityDialog dialog, NetOffice.OutlookSecurity.SecurityDialogCheckBox targetBox, NetOffice.OutlookSecurity.SecurityDialogLeftButton targetButton)
        {
            _dialog = dialog;
            _targetBox = targetBox;
            _targetButton = targetButton;
            if (this.InvokeRequired)
                this.Invoke(new MethodInvoker(Supress_OnAction));
            else
                Supress_OnAction();            

        }

        #endregion

        #region Methods

        private static string BoolToString(bool b)
        {
            if (b)
                return "true";
            else
                return "false";
        }

        private static Image ReadImageFromRessource(string ressourcePath)
        {
            System.IO.Stream ressourceStream = null;
            string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            ressourcePath = assemblyName + ".OutlookSecurity." + ressourcePath;
            ressourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(ressourcePath);
            if (ressourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));
            Bitmap newIcon = new Icon(ressourceStream).ToBitmap();
            return newIcon;
        }
        
        #endregion
    }
}
