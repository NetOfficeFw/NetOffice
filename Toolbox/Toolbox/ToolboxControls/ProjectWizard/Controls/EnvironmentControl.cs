using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Xml;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.Controls
{
    /// <summary>
    /// Step 2 in wizard to select programming language / ide / .net version
    /// </summary>
    [RessourceTable("ToolboxControls.ProjectWizard.Controls.EnvironmentControl.txt")]
    public partial class EnvironmentControl : UserControl, IWizardControl
    {
        #region Fields

        private XmlDocument _settings;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public EnvironmentControl()
        {
            InitializeComponent();
            CreateSettingsDocument();
            comboBoxNetRuntime.SelectedIndex = 0;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Selected language (C# or VB)
        /// </summary>
        internal string SelectedLanguage
        {
            get
            {
                if (radioButtonCSharp.Checked)
                    return "C#";
                else
                    return "VB";
            }
        }

        /// <summary>
        /// Selected IDE like VS2010
        /// </summary>
        internal string SelectedIDE
        {
            get
            {
                if (radioButtonVS2010.Checked)
                    return "2010";
                else
                    return "2013/2015/2017";
            }
        }

        /// <summary>
        /// Selected .NET Runtime
        /// </summary>
        internal string SelectedRuntime
        {
            get
            {
                return comboBoxNetRuntime.Text;
            }
        }

        #endregion

        #region IWizardControl

        public event ReadyStateChangedHandler ReadyStateChanged;
        
        public new void KeyDown(KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.D1:
                    radioButtonVB.Checked = true;
                    break;
                case Keys.D2:
                    radioButtonCSharp.Checked = true;
                    break;
            }
        }

        public bool IsReadyForNextStep
        {
            get { return true; }
        }

        public string Caption
        {
            get
            {
                return "Environment";
            }
        }

        public string Description
        {
            get
            {
                    return "The assistent prepare the implementation for you.";
            }
        }

        public ImageType Image
        {
            get
            {
                return ImageType.Question;
            }
        }

        public void Activate()
        {
             
        }

        public void Deactivate()
        { 
        
        }

        public System.Xml.XmlDocument SettingsDocument
        {
            get { return _settings; }
        }

        public string[] GetSettingsSummary()
        {
            
            string[] result = new string[2];
            result[0] += ProjectWizardControl.Singleton.Localized.Language + Environment.NewLine;
            result[1] += SelectedLanguage + Environment.NewLine;

            result[0] += "IDE" + Environment.NewLine;
            result[1] += "VS " + SelectedIDE + Environment.NewLine;

            result[0] += ProjectWizardControl.Singleton.Localized.Runtime;
            result[1] += SelectedRuntime;

            return result;
        }

        #endregion

        #region Mehtods

        private void ChangeSettings()
        {
            _settings.FirstChild.SelectSingleNode("Language").InnerText = SelectedLanguage;
            _settings.FirstChild.SelectSingleNode("IDE").InnerText = SelectedIDE;
            _settings.FirstChild.SelectSingleNode("Runtime").InnerText = SelectedRuntime;
        }

        private void CreateSettingsDocument()
        {
            _settings = new XmlDocument();
            _settings.AppendChild(_settings.CreateElement("EnvironmentControl"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("Language"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("IDE"));
            _settings.FirstChild.AppendChild(_settings.CreateElement("Runtime"));
        }

        private void RaiseChangeEvent()
        {
            if (null != ReadyStateChanged)
                ReadyStateChanged(this);
        }

        #endregion

        #region Trigger

        private void radioButtonLanguage_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                ChangeSettings();
                RaiseChangeEvent();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            } 
        }

        private void radioButtonIDE_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                ChangeSettings();
                RaiseChangeEvent();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            } 
        }

        private void comboBoxNetRuntime_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                ChangeSettings();
                RaiseChangeEvent();
                if (comboBoxNetRuntime.SelectedIndex >= 2)
                {
                    // .net 4.5
                    labelNet45Hint.Visible = true;
                    radioButtonVS2010.Enabled = false;
                    radioButtonVS2015.Enabled = true;
                    radioButtonVS2015.Checked = true;
                }
                else
                {
                    // else
                    labelNet45Hint.Visible = false;
                    radioButtonVS2010.Enabled = true;
                    radioButtonVS2015.Enabled = true;
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        #endregion
    }
}