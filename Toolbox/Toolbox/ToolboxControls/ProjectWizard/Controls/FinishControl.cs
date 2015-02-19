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
    public partial class FinishControl : UserControl, IWizardControl, ILocalizationDesign
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
                if (Forms.MainForm.Singleton.CurrentLanguageID == 1031)
                    return "Das Projekt wurde erfolgreich erstellt";
                else
                    return "The Project is successfully completed";
            }
        }

        public string Description
        {
            get
            {
                if (Forms.MainForm.Singleton.CurrentLanguageID == 1031)
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
            Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
            if (null != language)
            {
                var component = language.Components["Project Wizard - Finish"];
                Translation.Translator.TranslateControls(this, component.ControlRessources);
            }
            else
            {
                Translation.Translator.TranslateControls(this, "ToolboxControls.ProjectWizard.Controls.FinishControl.txt", Forms.MainForm.Singleton.CurrentLanguageID);
            }
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

        public IContainer Components
        {
            get { return components; }
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
                return new ILocalizationChildInfo[0];
            }
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
