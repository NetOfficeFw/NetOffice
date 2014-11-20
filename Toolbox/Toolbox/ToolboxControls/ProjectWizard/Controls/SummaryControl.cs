using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Xml;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.Controls
{
    [RessourceTable("ToolboxControls.ProjectWizard.Controls.SummaryControl.txt")]
    public partial class SummaryControl : UserControl, IWizardControl, ILocalizationDesign
    {
        private List<IWizardControl> _listControls;

        public SummaryControl()
        {
            InitializeComponent();
        }

        public SummaryControl(List<IWizardControl> listControls)
        {
            InitializeComponent();
            _listControls = listControls;
            //Translate();
        }

        #region IWizardControl 

        public new void KeyDown(KeyEventArgs e)
        {

        }

        public event ReadyStateChangedHandler ReadyStateChanged;

        public bool IsReadyForNextStep
        {
            get { return true; }
        }

        public string Caption
        {
            get
            {
                if (Forms.MainForm.Singleton.CurrentLanguageID == 1031)
                    return "Fertig! Nur noch ein Klick bis zur Erstellung.";
                else
                    return "Done!";
            }
        }

        public string Description
        {
            get
            {
                if (Forms.MainForm.Singleton.CurrentLanguageID == 1031)
                    return "Prüfen Sie Ihre Einstellungen in der Zusammenfassung.";
                else
                    return "Check your settings and lets get started.";
            }
        }

        public ImageType Image
        {
            get
            {
                return ImageType.Finish;
            }
        }

        public XmlDocument SettingsDocument
        {
            get
            {
                return null;
            }
        }

        public void Translate()
        {
            Translation.ToolLanguage language = Forms.MainForm.Singleton.Languages.Where(l => l.LCID == Forms.MainForm.Singleton.CurrentLanguageID).FirstOrDefault();
            if (null != language)
            {
                var component = language.Components["Project Wizard - Summary"];
                Translation.Translator.TranslateControls(this, component.ControlRessources);
            }
            else
            {
                Translation.Translator.TranslateControls(this, "ToolboxControls.ProjectWizard.Controls.SummaryControl.txt", Forms.MainForm.Singleton.CurrentLanguageID);
            }
            ShowSummary();
        }

        public void Activate()
        {
            RaiseChangeEvent();
            ShowSummary();
        }

        public void Deactivate()
        {

        }

        public string[] GetSettingsSummary()
        {
            string[] result = new string[2];
            result[0] = "";
            result[1] = "";
            return result;
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

        #region Methods

        private void ShowSummary()
        {
            labelSummaryCaption.Text = string.Empty;
            string summaryCaption = "";
            string summaryValue = "";

            foreach (Control item in _listControls)
            {                
                IWizardControl control = item as IWizardControl;
                if (null != control)
                {
                    if (ProjectWizardControl.Singleton.IsAddinProject)
                    {
                        string[] array = control.GetSettingsSummary();
                        summaryCaption += array[0] + Environment.NewLine;
                        summaryValue += array[1] + Environment.NewLine;
                    }
                    else
                    {
                        if ( (!(control is LoadControl)) && (!(control is GuiControl)))
                        {
                            string[] array = control.GetSettingsSummary();
                            summaryCaption += array[0] + Environment.NewLine;
                            summaryValue += array[1] + Environment.NewLine;
                        }
                    }                 
                }
            }
            labelSummaryCaption.Text = summaryCaption;
            labelSummaryValue.Text = summaryValue;
        }

        private void RaiseChangeEvent()
        {
            if (null != ReadyStateChanged)
                ReadyStateChanged(this);
        }

        #endregion
    }
}
