using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using Microsoft.Win32;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox
{
    public partial class ProjectWizardControl : UserControl, IToolboxControl
    {
        static int _currentLanguageID;
        static ProjectOptions _projectOptions;
        List<Control> _listControls = new List<Control>();
        Control _currentControl;

        public ProjectWizardControl()
        {
            InitializeComponent();
        }

        public static int CurrentLanguageID
        {
            get
            {
                return _currentLanguageID;
            }
        }

        public static ProjectOptions Options
        {
            get
            {
                return _projectOptions;
            }
        }

        #region IToolboxControl Member

        public string ControlName
        {
            get { return "ProjectWizard"; }
        }

        public string ControlCaption
        {
            get { return "VS Project Wizard"; }
        }

        public Image Icon
        {
            get
            {
                return ReadImageFromRessource("Icon.png");
            }
        }

        public void Activate()
        {
          
        }

        public void LoadComplete()
        {
           
        }

        public void LoadConfiguration(System.Xml.XmlNode configNode)
        {
        
        }

        public void SaveConfiguration(System.Xml.XmlNode configNode)
        {
            
        }

        public void SetLanguage(int id)
        {
            _currentLanguageID = id;
            Translator.TranslateControls(this, "ProjectWizard.MessageTable.txt", _currentLanguageID);
        }

        public IContainer Components
        {
            get { return components; }
        }

        #endregion

        #region Trigger

        private void buttonInfo_Click(object sender, EventArgs e)
         {
             try
             {
                 InfoControl infoBox = new InfoControl("ProjectWizard.Info" + _currentLanguageID.ToString() + ".rtf", true);
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

        private void buttonCreateProject_Click(object sender, EventArgs e)
        {
             CreateNewProject();
        }

        void currentControl_ReadyStateChanged(Control sender)
        {
            try
            {
                nextButton.Enabled = (sender as IWizardControl).IsReadyForNextStep;
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(exception);
            }
        }

        #endregion
         
        private void CreateNewProject()
        {
            SelectProjectTypeDialog dialog = new SelectProjectTypeDialog();
            if (DialogResult.OK == dialog.ShowDialog(this))
            {
                _projectOptions = dialog.SelectedOptions;
                LoadControls();
                SetActiveControl(_listControls[0]);
                panelHint.Visible = false;
                panelWizardHost.Visible = true;
            }
        }

        private void LoadControlsAddin()
        {
            HostControl control1 = new HostControl();
            NameControl control2 = new NameControl();
            LoadControl control3 = new LoadControl();
            GuiControl control4 = new GuiControl();
            _listControls.Add(control1);
            _listControls.Add(control2);
            _listControls.Add(control3);
            _listControls.Add(control4);

            panelControls.Controls.Add(control1);
            panelControls.Controls.Add(control2);
            panelControls.Controls.Add(control3);
            panelControls.Controls.Add(control4);

            control1.Dock = DockStyle.Fill;
            control2.Dock = DockStyle.Fill;
            control3.Dock = DockStyle.Fill;
            control4.Dock = DockStyle.Fill;

            SummaryControl control5 = new SummaryControl(_listControls);

            _listControls.Add(control5);
            panelControls.Controls.Add(control5);
            control5.Dock = DockStyle.Fill;
        }
     
        private void LoadControlsOther()
        {
            HostControl control1 = new HostControl();
            NameControl control2 = new NameControl();
            _listControls.Add(control1);
            _listControls.Add(control2);

            panelControls.Controls.Add(control1);
            panelControls.Controls.Add(control2);


            control1.Dock = DockStyle.Fill;
            control2.Dock = DockStyle.Fill;

            SummaryControl control5 = new SummaryControl(_listControls);

            _listControls.Add(control5);
            panelControls.Controls.Add(control5);
            control5.Dock = DockStyle.Fill;
        }

        private void LoadControls()
        {
            try
            {
                switch (_projectOptions.ProjectType)
                {
                    case ProjectType.Addin:
                        LoadControlsAddin();
                        break;
                    default:
                        LoadControlsOther();
                        break;
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("LoadControls " + ex.Message);
            }
        }

        private void Reset()
        {
            while (panelControls.Controls.Count > 0)
                panelControls.Controls.RemoveAt(0);
            _listControls.Clear();

            if (null != _currentControl)
            {
                IWizardControl control = _currentControl as IWizardControl;
                control.ReadyStateChanged -= new ReadyStateChangedHandler(currentControl_ReadyStateChanged);
                _currentControl = null;
            }
            panelHint.Visible = true;
            panelWizardHost.Visible = false;
        }
          
        private static Image ReadImageFromRessource(string ressourcePath)
        {
            System.IO.Stream ressourceStream = null;
            string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            ressourcePath = assemblyName + ".ProjectWizard." + ressourcePath;
            ressourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(ressourcePath);
            if (ressourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));
            Bitmap newIcon = new Bitmap(ressourceStream);
            return newIcon;
        }

        private int GetControlIndex(Control control)
        {
            int i = 0;
            foreach (IWizardControl item in _listControls)
            {
                if (item == control)
                    return i;
                i++;
            }
            throw new ArgumentOutOfRangeException("control");
        }

        private void GoToNextControl()
        {
            try
            {
                
                int currentIndex = GetControlIndex(_currentControl);
                Control control = _listControls[currentIndex + 1];
                SetActiveControl(control);
            }
            catch (Exception ex)
            {
                MessageBox.Show("GoToNextControl " + ex.Message);
            }
        }

        private void ReturnToPreviousControl()
        {
            try
            {
                int currentIndex = GetControlIndex(_currentControl);
                Control control = _listControls[currentIndex - 1];
                SetActiveControl(control);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ReturnToPreviousControl " + ex.Message);
            }
        }

        private bool IsLastControl(Control control)
        {
            return (_listControls[_listControls.Count - 1] == control);
        }

        private bool IsFirstControl(Control control)
        {
            return (_listControls[0] == control);
        }
         
        private void SetActiveControl(Control control)
        {
            try
            {
                foreach (Control item in panelControls.Controls)
                    item.Visible = false;

                control.Visible = true;
                _currentControl = control ;
                nextButton.Enabled = false;
                backButton.Enabled = !IsFirstControl(_currentControl);

                if (IsLastControl(_currentControl))
                {
                    finishButton.Location = nextButton.Location;
                    nextButton.Visible = false;
                    finishButton.Visible = true;
                }
                else
                {
                    nextButton.Visible = true;
                    finishButton.Visible = false;
                }

                (_currentControl as IWizardControl).Translate();
                (_currentControl as IWizardControl).Activate();
                labelCaption.Text = (_currentControl as IWizardControl).Caption;
                labelDescription.Text = (_currentControl as IWizardControl).Description;
                if ((_currentControl as IWizardControl).Image == ImageType.Question)
                    imageBox.Image = imageListIcons.Images[0];
                else
                    imageBox.Image = imageListIcons.Images[1];
                nextButton.Enabled = (_currentControl as IWizardControl).IsReadyForNextStep;

                if (CurrentLanguageID == 1031)
                    labelCurrentStep.Text = string.Format("Schritt {0} von {1}", GetControlIndex(_currentControl) + 1, _listControls.Count);
                else
                    labelCurrentStep.Text = string.Format("Step {0} of {1}", GetControlIndex(_currentControl) + 1, _listControls.Count);

                labelCurrentStep.Tag = new string[] { (GetControlIndex(_currentControl) + 1).ToString(), _listControls.Count.ToString() };

                (_currentControl as IWizardControl).ReadyStateChanged += new ReadyStateChangedHandler(currentControl_ReadyStateChanged);

            }
            catch (Exception ex)
            {
                MessageBox.Show("SetActiveControl " + ex.Message);
            }
        }

        private void backButton_Click(object sender, EventArgs e)
        {
            try
            {
                ReturnToPreviousControl();
                backButton.Focus();
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(exception);
            }
        }

        private void nextButton_Click(object sender, EventArgs e)
        {
            try
            {
                GoToNextControl();
                nextButton.Focus();
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(exception);
            }
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            Reset();
        }

        private void finishButton_Click(object sender, EventArgs e)
        {
            try
            {
                string message = CurrentLanguageID == 1031 ? "Das Projekt wurde erstellt." : "The project is complete.";
                
                string resultFolder = ProjectConverter.ConvertProjectTemplate(_projectOptions, _listControls);
                MessageBox.Show(this, message, "Developer Toolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);              
                System.Diagnostics.Process.Start(resultFolder);
                Reset();
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(exception);
            }
        }
    }
}
