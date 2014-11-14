using System;
using System.IO;
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
    public partial class FinishControl : UserControl, IWizardControl
    {
        private string _solutionPath;

        #region Ctor

        public FinishControl()
        {
            InitializeComponent();
        }

        #endregion

        #region Trigger

        private void FinishControl_Resize(object sender, EventArgs e)
        {
            panelHeader.Left = (animatedPanel1.Width / 2) - (panelHeader.Width / 2);
            panelButtons.Left = (animatedPanel1.Width / 2) - (panelButtons.Width / 2);            
        }

        #endregion

        #region Methods

        internal void SetSolutionPath(string solutionPath)
        {
            _solutionPath = solutionPath;
            bool fileExists = File.Exists(_solutionPath);
            buttonOpenSolution.Enabled = fileExists;
            buttonOpenFolder.Enabled = fileExists;
        }

        #endregion

        #region IWizardControl

        public event ReadyStateChangedHandler ReadyStateChanged;

        private void RaiseReadyStateChanged()
        {
            if (null != ReadyStateChanged)
                ReadyStateChanged(this);
        }

        public bool IsReadyForNextStep
        {
            get { return true; }
        }

        public string Caption
        {
            get 
            {
                if (ProjectWizardControl.Singleton.Host.CurrentLanguageID == 1031)
                    return "Das Projekt wurde erfolgreich erstellt";
                else
                    return "The Project is successfully completed";
            }
        }

        public string Description
        {
            get
            {
                if (ProjectWizardControl.Singleton.Host.CurrentLanguageID == 1031)
                    return "Viel Erfolg bei der Arbeit.";
                else
                    return "We wish you much success in your work";
            }
        }

        public ImageType Image
        {
            get { return ImageType.Finish; } 
        }

        public void Translate()
        {
            Translation.Translator.TranslateControls(this, "ToolboxControls.ProjectWizard.Controls.FinishControl.txt", ProjectWizardControl.Singleton.Host.CurrentLanguageID);
        }

        public void Activate()
        {
            animatedPanel1.Animation1Enabled = true;
            controlBackColorAnimator1.Start(false);
            controlForeColorAnimator1.Start(false);
        }

        public void Deactivate()
        {
            animatedPanel1.Animation1Enabled = false;
            controlBackColorAnimator1.Stop();
            controlForeColorAnimator1.Stop();
        }

        public XmlDocument SettingsDocument
        {
            get { throw new NotImplementedException(); }
        }

        public string[] GetSettingsSummary()
        {
            throw new NotImplementedException();
        }

        public new void KeyDown(KeyEventArgs e)
        {
            
        }

        #endregion

        #region Trigger

        private void buttonOpenSolution_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(_solutionPath);
                buttonClose_Click(buttonClose, EventArgs.Empty);
            }
            catch
            {
                ;
            }
        }

        private void buttonOpenFolder_Click(object sender, EventArgs e)
        {
            try
            {
                string folderPath = Path.GetDirectoryName(_solutionPath);
                System.Diagnostics.Process.Start(folderPath);
                buttonClose_Click(buttonClose, EventArgs.Empty);
            }
            catch
            {
                ;
            }
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            try
            {
                RaiseReadyStateChanged();
            }
            catch
            {
                ;
            }
        }

        #endregion
    }
}
