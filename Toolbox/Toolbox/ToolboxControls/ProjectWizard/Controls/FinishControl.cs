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
    /// <summary>
    /// Last wizard step to show summary settings
    /// </summary>
    [RessourceTable("ToolboxControls.ProjectWizard.Controls.FinishControl.txt")]
    public partial class FinishControl : UserControl, IWizardControl
    {
        #region Fields

        private string _solutionPath;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public FinishControl()
        {
            InitializeComponent();
        }

        #endregion

        #region Methods

        /// <summary>
        /// Set target solution path
        /// </summary>
        /// <param name="solutionPath">target solution path</param>
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
                return "The Project is successfully completed";
            }
        }

        public string Description
        {
            get
            {
                    return "We wish you much success in your work";
            }
        }

        public ImageType Image
        {
            get { return ImageType.Finish; } 
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


        private void FinishControl_Resize(object sender, EventArgs e)
        {
            try
            {
                panelHeader.Left = (animatedPanel1.Width / 2) - (panelHeader.Width / 2);
                panelButtons.Left = (animatedPanel1.Width / 2) - (panelButtons.Width / 2);
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            } 
        }

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
