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

        List<IWizardControl> _listControls = new List<IWizardControl>();
        IWizardControl _currentControl;

        public ProjectWizardControl()
        {
            InitializeComponent();
            Singleton = this;
        }

        public static int CurrentLanguageID
        {
            get
            {
                return _currentLanguageID;
            }
        }
         
        #region IToolboxControl Member

        public new void KeyDown(KeyEventArgs e)
        {
            foreach (var item in _listControls)
	        {
                Control winControl = item as Control;
                if(winControl.Visible)
    		        item.KeyDown(e);
	        }
        }

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
 
         
        private void CreateNewProject()
        {
            Reset();
            LoadControls();
            SetActiveControl(_listControls[0]);
            panelHint.Visible = false;
            panelWizardHost.Visible = true;
        }

        internal static ProjectWizardControl Singleton { get; private set; }

        private void LoadControls()
        {
            ProjectControl control0 = new ProjectControl();
            EnvironmentControl control1 = new EnvironmentControl();
            HostControl control2 = new HostControl();
            NameControl control3 = new NameControl();
            LoadControl control4 = new LoadControl();
            GuiControl control5 = new GuiControl();

            _listControls.Add(control0);
            _listControls.Add(control1);
            _listControls.Add(control2);
            _listControls.Add(control3);
            _listControls.Add(control4);
            _listControls.Add(control5);

            panelControls.Controls.Add(control0);
            panelControls.Controls.Add(control1);
            panelControls.Controls.Add(control2);
            panelControls.Controls.Add(control3);
            panelControls.Controls.Add(control4);
            panelControls.Controls.Add(control5);

            control0.Dock = DockStyle.Fill;
            control1.Dock = DockStyle.Fill;
            control2.Dock = DockStyle.Fill;
            control3.Dock = DockStyle.Fill;
            control4.Dock = DockStyle.Fill;
            control4.Dock = DockStyle.Fill;

            SummaryControl control6 = new SummaryControl(_listControls);

            _listControls.Add(control6);
            panelControls.Controls.Add(control6);
            control6.Dock = DockStyle.Fill;
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

        private int GetControlIndex(IWizardControl control)
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

        internal bool FolderExists(string name)
        {

            foreach (var item in _listControls)
            {
                ProjectControl ctrl = item as ProjectControl;
                if (null != ctrl)
                {
                    string basePath = ctrl.CalculatedFolder;
                    string fullPath = System.IO.Path.Combine(basePath, name);
                    return System.IO.Directory.Exists(fullPath) || System.IO.File.Exists(fullPath);
                }
            }
            return false;
        }

        internal bool IsSingleMSProjectProject
        {
            get
            {
                bool visioChecked = false;
                bool otherChecked = false;
                foreach (var item in _listControls)
                {
                    HostControl ctrl = item as HostControl;
                    if (null != ctrl)
                    {
                        foreach (Control winCtrl in (ctrl as Control).Controls)
	                    {
                            CheckBox box = winCtrl as CheckBox;
                            if (null != box)
                            {
                                if (box.Name == "checkBoxProject" && box.Checked)
                                {
                                    visioChecked = true;
                                }
                                else if (box.Checked)
                                {
                                    otherChecked = true;
                                }
                            }
	                    }   
                    }
                }

                return visioChecked == true || otherChecked == false;
            }
        }

        internal bool IsAddinProject
        {
            get
            {
                foreach (var item in _listControls)
                {
                    ProjectControl ctrl = item as ProjectControl;
                    if (null != ctrl)
                    {
                        if ("AutomationAddin" == ctrl.SelectedProjectType(1033))
                            return true;
                    }
                }
                return false;
            }
        }

        private void GoToNextControl()
        {
            try
            {
                int currentIndex = GetControlIndex(_currentControl);
                if (!IsAddinProject)
                {
                    if (_listControls[currentIndex + 1] is LoadControl)
                    {
                        IWizardControl control = _listControls[currentIndex + 3];
                        SetActiveControl(control);
                    }
                    else
                    {
                        IWizardControl control = _listControls[currentIndex + 1];
                        SetActiveControl(control);
                    }
                }
                else
                {
                    IWizardControl control = _listControls[currentIndex + 1];
                    SetActiveControl(control);
                }
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
                if (!IsAddinProject)
                {
                    if (_listControls[currentIndex - 1] is LoadControl)
                    {
                        IWizardControl control = _listControls[currentIndex - 3];
                        SetActiveControl(control);
                    }
                    else
                    {
                        IWizardControl control = _listControls[currentIndex - 1];
                        SetActiveControl(control);
                    }
                }
                else
                {
                    IWizardControl control = _listControls[currentIndex - 1];
                    SetActiveControl(control);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ReturnToPreviousControl " + ex.Message);
            }
        }

        private bool IsLastControl(IWizardControl control)
        {
            return (_listControls[_listControls.Count - 1] == control);
        }

        private bool IsFirstControl(IWizardControl control)
        {
            return (_listControls[0] == control);
        }
         
        #region Trigger

        private void buttonInfo_Click(object sender, EventArgs e)
        {
            try
            {
                InfoControl infoBox = new InfoControl("ProjectWizard.ProjectWizard.Info." + _currentLanguageID.ToString() + ".rtf", true);
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
                ErrorForm.ShowError(this, exception);
            }
        }

        private void SetActiveControl(IWizardControl control)
        {
            try
            {
                foreach (Control item in panelControls.Controls)
                    item.Visible = false;

                (control as Control).Visible = true;
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
                    labelCurrentStep.Text = string.Format("Schritt {0} von {1}", GetControlIndex(_currentControl) + 1, IsAddinProject == true ? _listControls.Count : _listControls.Count - 2);
                else
                    labelCurrentStep.Text = string.Format("Step {0} of {1}", GetControlIndex(_currentControl) + 1, IsAddinProject == true ? _listControls.Count : _listControls.Count - 2);

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
                ErrorForm.ShowError(this, exception);
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
                ErrorForm.ShowError(this, exception);
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
                string resultFolder = ProjectConverter.ConvertProjectTemplate(_listControls);
                FinishDialog dialog = new FinishDialog(resultFolder);
                dialog.ShowDialog(this);
                Reset();
            }
            catch (Exception exception)
            {
                ErrorForm.ShowError(this, exception);
            }
        }

        #endregion

        #region Static Methods

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

        #endregion
    }
}
