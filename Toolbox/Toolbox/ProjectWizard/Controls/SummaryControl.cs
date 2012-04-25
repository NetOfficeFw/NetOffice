using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Xml;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox
{
    public partial class SummaryControl : UserControl, IWizardControl
    {
        List<Control> _listControls;

        public SummaryControl(List<Control> listControls)
        {
            InitializeComponent();
            _listControls = listControls;
            Translate();
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
                if (ProjectWizardControl.CurrentLanguageID == 1031)
                    return "Fertig! Nur noch ein Klick bis zur Erstellung.";
                else
                    return "Done!";
            }
        }

        public string Description
        {
            get
            {
                if (ProjectWizardControl.CurrentLanguageID == 1031)
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
            if (ProjectWizardControl.CurrentLanguageID == 1031)
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

            labelSummaryCaption.Text = string.Empty;
            string summaryCaption = "";
            string summaryValue = "";

            foreach (Control item in _listControls)
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
