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
    /// <summary>
    /// Show all selected options as summary
    /// </summary>
    [RessourceTable("ToolboxControls.ProjectWizard.Controls.SummaryControl.txt")]
    public partial class SummaryControl : UserControl, IWizardControl
    {
        #region Fields

        private List<IWizardControl> _listControls;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public SummaryControl()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="listControls">all used wizard steps with options</param>
        public SummaryControl(List<IWizardControl> listControls)
        {
            InitializeComponent();
            _listControls = listControls;
        }

        #endregion

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
                return "Done!";
            }
        }

        public string Description
        {
            get
            {
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