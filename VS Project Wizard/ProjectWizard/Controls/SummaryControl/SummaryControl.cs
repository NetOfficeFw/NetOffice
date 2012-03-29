using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Xml;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.ProjectWizard
{
    internal partial class SummaryControl : UserControl, IWizardControl
    {
        NetOfficeProject _currentProject;

        public SummaryControl(NetOfficeProject currentProject)
        {
            InitializeComponent();

            Translate();

            _currentProject = currentProject;
        }

        #region IWizardControl Member

        public event ReadyStateChangedHandler ReadyStateChanged;

        public bool IsReadyForNextStep
        {
            get { return true; }
        }

        public string Caption
        {
            get
            {
                if (NetOfficeProject.TargetLanguage == TargetLanguage.German)
                    return "Fertig! Nur noch ein Klick bis zur Erstellung.";
                else
                    return "Done!";
            }
        }

        public string Description
        {
            get
            {
                if (NetOfficeProject.TargetLanguage == TargetLanguage.German)
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
            if (NetOfficeProject.TargetLanguage == TargetLanguage.German)
                labelSummaryHeader.Text = "Ausgwählte Einstellungen in der Übersicht";
            else
                labelSummaryHeader.Text = "Summary Table";

            ShowSummary();
        }

        public void Activate()
        {
            RaiseChangeEvent();
            ShowSummary();
        }

        public string[] GetSettingsSummary()
        {
            string[] result = new string[2];
            result[0] = "";
            result[1] = "";
            return result;
        }

        #endregion

        #region Methods

        private void ShowSummary()
        {
            if (null == _currentProject)
                return;

           labelSummaryCaption.Text = string.Empty;
           string summaryCaption = "";
           string summaryValue  = "";  
            
           foreach (Control item in _currentProject.ListControls)
	       {
               IWizardControl control = item as IWizardControl;
               if (null != control)
               {
                   string[] array = control.GetSettingsSummary();
                   summaryCaption += array[0] + Environment.NewLine;
                   summaryValue += array[1] + Environment.NewLine;
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
