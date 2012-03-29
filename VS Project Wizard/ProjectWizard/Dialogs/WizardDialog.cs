using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Xml;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace NetOffice.ProjectWizard
{
    partial class WizardDialog : Form
    {
        #region Fields

        IWizardControl _currentControl;
        static NetOfficeProject _parent;

        bool languageChangeFlag;

        #endregion

        #region Construction

        public WizardDialog(NetOfficeProject parent)
        {
            try
            {
                this.Tag = parent.Name;

                _parent = parent;
                InitializeComponent();
                LoadControls();

                if (TargetLanguage.German == NetOfficeProject.TargetLanguage)
                    this.Text = string.Format("Assistent für {0}", parent.Name);
                else
                    this.Text = string.Format("{0} Assistant", parent.Name);

                Translator.TranslateControls(this, "Dialogs.WizardDialog.txt", NetOfficeProject.TargetLanguage, true);
             
                languageChangeFlag = true;
                
                if (TargetLanguage.German == NetOfficeProject.TargetLanguage)
                    comboBoxLanguage.SelectedIndex = 1;
                else
                    comboBoxLanguage.SelectedIndex = 0;

                languageChangeFlag = false;

                if (NetOfficeProject.TargetLanguage == TargetLanguage.German)
                    labelCurrentStep.Text = string.Format("Schritt {0} von {1}", GetControlIndex(_currentControl) + 1, _parent.ListControls.Count);
                else
                    labelCurrentStep.Text = string.Format("Step {0} of {1}", GetControlIndex(_currentControl) + 1, _parent.ListControls.Count);

                labelCurrentStep.Tag = new string[] { (GetControlIndex(_currentControl) + 1).ToString(), _parent.ListControls.Count.ToString() };

            }
            catch (Exception exception)
            {
                ErrorDialog dialog = new ErrorDialog(exception, NetOfficeProject.TargetLanguage);
                dialog.ShowDialog();
            }
        }
        
        #endregion

        #region Methods

        private int GetControlIndex(IWizardControl control)
        {
            int i=0;
            foreach (IWizardControl item in _parent.ListControls)
            {
                if (item == control)
                    return i;
                i++;
            }
            throw new ArgumentOutOfRangeException("control");
        }

        private bool IsLastControl(IWizardControl control)
        {
            return (_parent.ListControls[_parent.ListControls.Count - 1] == control);
        }

        private bool IsFirstControl(IWizardControl control)
        {
            return (_parent.ListControls[0] == control);
        }

        private void LoadControls()
        {
            try
            {
                foreach (Control item in _parent.ListControls)
                    SetControlToPanel(item);
                 
                SetActiveControl(_parent.ListControls[0]);
            }
            catch (Exception ex)
            {
                MessageBox.Show("LoadControls " + ex.Message);
            }
        }

        private void SetActiveControl(Control control)
        {
            try
            {
                foreach (Control item in panelControls.Controls)
                    item.Visible = false;

                control.Visible = true;
                _currentControl = control as IWizardControl;
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

                _currentControl.Translate();
                _currentControl.Activate();
                labelCaption.Text = _currentControl.Caption;
                labelDescription.Text = _currentControl.Description; 
                if (_currentControl.Image == ImageType.Question)
                    imageBox.Image = imageListIcons.Images[0];
                else
                    imageBox.Image = imageListIcons.Images[1];
                nextButton.Enabled = _currentControl.IsReadyForNextStep;

                if (NetOfficeProject.TargetLanguage == TargetLanguage.German)
                    labelCurrentStep.Text = string.Format("Schritt {0} von {1}", GetControlIndex(_currentControl) + 1, _parent.ListControls.Count);
                else
                    labelCurrentStep.Text = string.Format("Step {0} of {1}", GetControlIndex(_currentControl) + 1, _parent.ListControls.Count);

                labelCurrentStep.Tag = new string[] { (GetControlIndex(_currentControl) + 1).ToString(), _parent.ListControls.Count.ToString()};

            }
            catch (Exception ex)
            {
                MessageBox.Show("SetActiveControl " + ex.Message);
            }
        }

        private void SetControlToPanel(Control control)
        {
            try
            {
                panelControls.Controls.Add(control);
                control.Dock = DockStyle.Fill;
                (control as IWizardControl).ReadyStateChanged += new ReadyStateChangedHandler(currentControl_ReadyStateChanged);

            }
            catch (Exception ex)
            {
                MessageBox.Show("SetControlToPanel " + ex.Message);
            }
        }
       
        private void GoToNextControl()
        {
            try
            {
                int currentIndex = GetControlIndex(_currentControl);
                Control control = _parent.ListControls[currentIndex + 1];
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
                Control control = _parent.ListControls[currentIndex - 1];
                SetActiveControl(control);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ReturnToPreviousControl " + ex.Message);
            }
        }

        #endregion

        #region Trigger
 
        private void backButton_Click(object sender, EventArgs e)
        {
            try
            {
                ReturnToPreviousControl();
                backButton.Focus();
            }
            catch (Exception exception)
            {
                ErrorDialog dialog = new ErrorDialog(exception, NetOfficeProject.TargetLanguage);
                dialog.ShowDialog();
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
                ErrorDialog dialog = new ErrorDialog(exception, NetOfficeProject.TargetLanguage);
                dialog.ShowDialog();
            }
        }

        private void finishButton_Click(object sender, EventArgs e)
        {
            try
            {
                _parent.FinishAction();

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception exception)
            {
                ErrorDialog dialog = new ErrorDialog(exception, NetOfficeProject.TargetLanguage);
                dialog.ShowDialog();
            }            
        }
      
        void currentControl_ReadyStateChanged(IWizardControl sender)
        {
            try
            {
                nextButton.Enabled = sender.IsReadyForNextStep;
            }
            catch (Exception exception)
            {
                ErrorDialog dialog = new ErrorDialog(exception, NetOfficeProject.TargetLanguage);
                dialog.ShowDialog();
            }
        }

        private void comboBoxLanguage_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (languageChangeFlag)
                    return;

                if (0 == comboBoxLanguage.SelectedIndex)
                    NetOfficeProject.TargetLanguage = TargetLanguage.English;
                else
                    NetOfficeProject.TargetLanguage = TargetLanguage.German;

                Translator.TranslateControls(this, "Dialogs.WizardDialog.txt", NetOfficeProject.TargetLanguage, true);

                labelCaption.Text = _currentControl.Caption;
                labelDescription.Text = _currentControl.Description;
                _currentControl.Translate();
                
                if (TargetLanguage.German == NetOfficeProject.TargetLanguage)
                {
                    this.Text = string.Format("Assistent für {0}", _parent.Name);
                    labelCurrentStep.Text = string.Format("Schritt {0} von {1}", GetControlIndex(_currentControl) + 1, _parent.ListControls.Count);
                }
                else
                {
                    this.Text = string.Format("{0} Assistant", _parent.Name);
                    labelCurrentStep.Text = string.Format("Step {0} of {1}", GetControlIndex(_currentControl) + 1, _parent.ListControls.Count);
                }
                 
                labelCurrentStep.Tag = new string[] { (GetControlIndex(_currentControl) + 1).ToString(), _parent.ListControls.Count.ToString() };

            }
            catch (Exception exception)
            {
                ErrorDialog dialog = new ErrorDialog(exception, NetOfficeProject.TargetLanguage);
                dialog.ShowDialog();
            }
        }

        #endregion
    }
}
