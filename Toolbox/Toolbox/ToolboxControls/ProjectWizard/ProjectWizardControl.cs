using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard.Controls;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProjectWizard
{
    /// <summary>
    /// Allows to create new development projects in c# or vb
    /// </summary>
    [RessourceTable("ToolboxControls.ProjectWizard.Strings.txt")]
    public partial class ProjectWizardControl : UserControl, IToolboxControl
    {
        #region Fields

        private List<IWizardControl> _listControls = new List<IWizardControl>();
        private IWizardControl _currentControl;
        private FinishControl _finishControl;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public ProjectWizardControl()
        {
            InitializeComponent();
            Localized = new LocalizedContent();
            Captions = new LocalizedCaptions();
            Singleton = this;
        }

        #endregion

        #region Properties

        /// <summary>
        /// All wizard steps
        /// </summary>
        public List<IWizardControl> WizardControls
        {
            get
            {
                return _listControls;
            }
        }

        /// <summary>
        /// Singleton to made easy access for wizard controls
        /// </summary>
        internal static ProjectWizardControl Singleton { get; private set; }

        /// <summary>
        /// Wizard has been started
        /// </summary>
        private bool IsCurrentlyActive { get { return panelWizardHost.Visible; } }

        /// <summary>
        /// Localized messages
        /// </summary>
        internal LocalizedContent Localized { get; private set; }

        /// <summary>
        /// Localized captions
        /// </summary>
        internal LocalizedCaptions Captions { get; private set; }

        #endregion

        #region IToolboxControl

        public IToolboxHost Host { get; private set; }

        public new void KeyDown(KeyEventArgs e)
        {
            if (IsCurrentlyActive)
            {
                if (e.KeyCode == Keys.Return && null != _currentControl && true == nextButton.Enabled && (!(_currentControl is SummaryControl)))
                {
                   nextButton_Click(nextButton, EventArgs.Empty);
                }
                else if (e.KeyCode == Keys.Return && null != _currentControl && true == nextButton.Enabled && (_currentControl is SummaryControl))
                {
                    finishButton_Click(finishButton, EventArgs.Empty);
                }
                else
                {
                    foreach (var item in _listControls)
                    {
                        Control winControl = item as Control;
                        if (winControl.Visible)
                            item.KeyDown(e);
                    }
                }
            }
            else
            {
                if (e.KeyCode == Keys.Return)
                    buttonCreateProject_Click(buttonCreateProject, EventArgs.Empty);
            }
        }

        public string ControlName
        {
            get { return "ProjectWizard.ProjectWizardControl"; }
        }

        public string ControlCaption
        {
            get { return "Project Wizard"; }
        }

        public Image Icon
        {
            get { return Ressources.RessourceUtils.ReadImageFromRessource("ToolboxControls.ProjectWizard.Icon.png"); }
        }

        public bool SupportsInfoMessage
        {
            get
            {
                return false;
            }
        }

        public ToolboxControlMessageKind InfoMessageKind
        {
            get
            {
                return ToolboxControlMessageKind.Uncategorized;
            }
        }

        public string InfoMessage
        {
            get
            {
                return String.Empty;
            }
        }

        public bool SupportsHelpContent
        {
            get
            {
                return true;
            }
        }

        public void InitializeControl(IToolboxHost host)
        {
            Host = host;
        }

        public void Activate(bool firstTime)
        {
            controlForeColorAnimator1.Start(false);
            controlBackColorAnimator1.Start(false);
        }

        public void Deactivated()
        {
            controlForeColorAnimator1.Stop();
            controlBackColorAnimator1.Stop();

            if (IsCurrentlyActive)
                Reset();
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

        public Stream GetHelpText()
        {
            return Ressources.RessourceUtils.ReadStream("ToolboxControls.ProjectWizard.Info1033.rtf");
        }

        public void Release()
        {

        }

        public IContainer Components
        {
            get { return components; }
        }

        #endregion

        #region Methods

        private void CreateNewProject()
        {
            Reset();
            LoadControls();
            SetActiveControl(_listControls[0]);
            panelHint.Visible = false;
            panelWizardHost.Visible = true;
        }

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
            control5.Dock = DockStyle.Fill;

            SummaryControl control6 = new SummaryControl(_listControls);

            _listControls.Add(control6);
            panelControls.Controls.Add(control6);
            control6.Dock = DockStyle.Fill;

            _finishControl = new FinishControl();
            panelControls.Controls.Add(_finishControl);
            _finishControl.Dock = DockStyle.Fill;
            _finishControl.ReadyStateChanged += new ReadyStateChangedHandler(FinishControl_ReadyStateChanged);
        }

        private void Reset()
        {
            for (int i = 1; i <= 7; i++)
            {
                string controlName = String.Format("pictureBoxStep{0}", i);
                PictureBox controlBox = panelLeftHeader.Controls[controlName] as PictureBox;
                controlBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            }

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
    
            if(null != _finishControl)
                _finishControl.Deactivate();
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

        internal bool IsSingleVisioProject
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
                                if (box.Name == "checkBoxVisio" && box.Checked)
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
                        if (ProjectType.NetOfficeAddin == ctrl.SelectedProjectType() || ProjectType.SimpleAddin == ctrl.SelectedProjectType())
                            return true;
                    }
                }
                return false;
            }
        }

        internal bool IsSimpleAddinProject
        {
            get
            {
                foreach (var item in _listControls)
                {
                    ProjectControl ctrl = item as ProjectControl;
                    if (null != ctrl)
                    {
                        if (ProjectType.SimpleAddin == ctrl.SelectedProjectType())
                            return true;
                    }
                }
                return false;
            }
        }

        private void BorderCurrentStep(int step)
        {
            for (int i = 1; i <= 7; i++)
            {
                string controlName = String.Format("pictureBoxStep{0}", i);
                PictureBox controlBox = panelLeftHeader.Controls[controlName] as PictureBox;
                controlBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            }

            string targetControlName = String.Format("pictureBoxStep{0}", step + 1);
            PictureBox stepBox = panelLeftHeader.Controls[targetControlName] as PictureBox;
            if (null != stepBox)
                stepBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
        }

        private void SetActiveControl(IWizardControl control)
        {
            try
            {
                IWizardControl oldActiveControl = _currentControl as IWizardControl;
                if (null != oldActiveControl)
                    oldActiveControl.ReadyStateChanged -= new ReadyStateChangedHandler(currentControl_ReadyStateChanged);

                foreach (Control item in panelControls.Controls)
                    item.Visible = false;

                if (null != oldActiveControl)
                    oldActiveControl.Deactivate();

                (control as Control).Visible = true;
                _currentControl = control;
                nextButton.Enabled = false;
                backButton.Enabled = !IsFirstControl(_currentControl);
                backButton.Visible = true;
                cancelButton.Visible = true;
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
                
                (_currentControl as IWizardControl).Activate();
                labelCaption.Text = Captions.GetCaption(_currentControl as IWizardControl);
                labelDescription.Text = Captions.GetDescription(_currentControl as IWizardControl);
                if ((_currentControl as IWizardControl).Image == ImageType.Question)
                    imageBox.Image = imageListIcons.Images[0];
                else
                    imageBox.Image = imageListIcons.Images[1];
                nextButton.Enabled = (_currentControl as IWizardControl).IsReadyForNextStep;

                int currentIndex = GetControlIndex(_currentControl) + 1;
                int totalCount = IsAddinProject == true ? _listControls.Count : _listControls.Count - 2;
                if (currentIndex > totalCount)
                    currentIndex = totalCount;

                labelCurrentStep.Text = Localized.StepProgress.Replace("{0}", currentIndex.ToString()).Replace("{1}", totalCount.ToString());

                labelCurrentStep.Tag = new string[] { (GetControlIndex(_currentControl) + 1).ToString(), _listControls.Count.ToString() };

                (_currentControl as IWizardControl).ReadyStateChanged += new ReadyStateChangedHandler(currentControl_ReadyStateChanged);

                int index = GetControlIndex(_currentControl);
                BorderCurrentStep(index);
                (_currentControl as Control).Focus();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception);
            }
        }

        private void ShowFinish(string solutionPath)
        {
            IWizardControl oldActiveControl = _currentControl as IWizardControl;
            if (null != oldActiveControl)
                oldActiveControl.ReadyStateChanged -= new ReadyStateChangedHandler(currentControl_ReadyStateChanged);

            foreach (Control item in panelControls.Controls)
                item.Visible = false;

            backButton.Visible = false;
            nextButton.Visible = false;
            finishButton.Visible = false;
            cancelButton.Visible = false;

            _finishControl.Activate();
            
            labelCurrentStep.Text = Localized.Completed;;
            labelCaption.Text = _finishControl.Caption;
            labelDescription.Text = _finishControl.Description;
            if (_finishControl.Image == ImageType.Question)
                imageBox.Image = imageListIcons.Images[0];
            else
                imageBox.Image = imageListIcons.Images[1];

            _finishControl.SetSolutionPath(solutionPath);

            labelCaption.Text = _finishControl.Caption;
            _finishControl.Visible = true;
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
                    else if (_listControls[currentIndex] is SummaryControl)
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

        #endregion

        #region Trigger

        private void FinishControl_ReadyStateChanged(Control sender)
        {
            if (_finishControl.IsReadyForNextStep)
                Reset();
        }

        private void buttonCreateProject_Click(object sender, EventArgs e)
        {
            try
            {
                CreateNewProject();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception);
            }
        }

        private void currentControl_ReadyStateChanged(Control sender)
        {
            try
            {
                nextButton.Enabled = (sender as IWizardControl).IsReadyForNextStep;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception);
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
                Forms.ErrorForm.ShowError(this, exception);
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
                Forms.ErrorForm.ShowError(this, exception);
            }
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            try
            {
                Reset();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception);
            }
        }

        private void finishButton_Click(object sender, EventArgs e)
        {
            try
            {
                ProjectConverters.Converter converter = ProjectConverters.Converter.CreateConverter(new ProjectOptions(_listControls));
                string solutionPath = converter.CreateSolution();
                converter.Dispose();
                ShowFinish(solutionPath);
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception);
            }
        }

        private void ProjectWizardControl_Resize(object sender, EventArgs e)
        {
            try
            {
                panelHint.Location = new Point((this.Width / 2) - (panelHint.Width / 2)+2, (this.Height / 2) - (panelHint.Height / 2) -59);
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception);
            }
        }

        #endregion
    }
}
