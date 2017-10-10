using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace TutorialsBase
{
    public partial class TutorialForm : Form, IHost
    {
        #region Ctor

        public TutorialForm()
        {
            InitializeComponent();
            Tutorials = new BindingList<ITutorial>();
            ExamplesView.AutoGenerateColumns = false;
            ExamplesView.DataSource = Tutorials;
        }

        #endregion
        #region Properties

        private BindingList<ITutorial> Tutorials { get; set; }

        private ITutorial SelectedTutorial
        {
            get
            {
                if (ExamplesView.SelectedRows.Count > 0)
                    return ExamplesView.SelectedRows[0].DataBoundItem as ITutorial;
                else
                    return null;
            }
        }

        #endregion

        #region Methods

        protected internal void LoadTutorial(ITutorial tutorial)
        {
            tutorial.Connect(this);
            Tutorials.Add(tutorial);
        }

        #endregion

        #region IHost

        public DialogResult ShowQuestion(string message)
        {
            return MessageBox.Show(this, message, "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        }

        public void ShowMessage(string message)
        {
            try
            {
                MessageBox.Show(this, message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception exception)
            {
                ErrorForm.Show(this, null, exception.Message, exception);
            }
        }

        public void ShowFinishDialog(string message)
        {
            try
            {
                FinishForm dialog = new FinishForm(message);
                dialog.ShowDialog(this);
            }
            catch(Exception exception)
            {
                ErrorForm.Show(this, null, exception.Message, exception);               
            }
        }

        public void ShowFinishDialog()
        {
            try
            {
                FinishForm dialog = new FinishForm();
                dialog.ShowDialog(this);
            }
            catch (Exception exception)
            {
                ErrorForm.Show(this, null, exception.Message, exception);
            }
        }

        public void ShowErrorDialog(string message, Exception exception)
        {
            ErrorForm.Show(this, null, message, exception);
        }

        public Icon DisplayIcon
        {
            get 
            {
                return this.Icon;
            }
        }

        #endregion

 
        #region UI Trigger

        private void ExamplesView_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                if (null != SelectedTutorial)
                    AreaForm.ShowForm(this, SelectedTutorial);
            }
            catch (Exception exception)
            {
                ErrorForm.Show(this, null, exception.Message, exception);
            }
        }

        private void buttonOptions_Click(object sender, EventArgs e)
        {
            try
            {
                OptionsForm dialog = new OptionsForm();
                dialog.ShowDialog(this);
                dialog.Dispose();
            }
            catch(Exception exception)
            {
                ErrorForm.Show(this, null, exception.Message, exception);  
            }
        }

        private void linkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                LinkLabel label = sender as LinkLabel;
                string link = label.Tag as string;
                System.Diagnostics.Process.Start(link);
            }
            catch (Exception exception)
            {
                ErrorForm.Show(this, null, exception.Message, exception);
            }
        }

        private void FormBase_FormClosed(object sender, FormClosedEventArgs e)
        {
            foreach (ITutorial item in Tutorials)
            {
                try
                {
                    item.Disconnect();
                }
                catch (Exception exception)
                {
                    Console.WriteLine("ITutorial Disconnect Exception:" + exception.Message);
                }
            }
        }

        #endregion
    }
}
