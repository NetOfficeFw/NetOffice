using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace TutorialsBase
{
    /// <summary>
    /// the main form for the application, implements also IHost to represent the host application for the examples.
    /// this is not -state of the art- in software development but example code has to keep the lines of code as small as possible.
    /// </summary>
    public partial class FormBase : TutorialForm, IHost
    {
        #region Fields

        List<ITutorial> _listTutorials = new List<ITutorial>();

        #endregion

        #region Ctor

        public FormBase()
        {
            InitializeComponent();
            FormOptions.LoadConfigurationFromXMLFile(this);
            ApplyConfiguration();
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
                FormError.Show(this, null, exception.Message, exception);
            }
        }

        public void ShowFinishDialog(string message)
        {
            try
            {
                FormFinish dialog = new FormFinish(message);
                dialog.ShowDialog(this);
            }
            catch(Exception exception)
            {
                FormError.Show(this, null, exception.Message, exception);               
            }
        }

        public void ShowFinishDialog()
        {
            try
            {
                FormFinish dialog = new FormFinish();
                dialog.ShowDialog(this);
            }
            catch (Exception exception)
            {
                FormError.Show(this, null, exception.Message, exception);
            }
        }

        public void ShowErrorDialog(string message, Exception exception)
        {
            FormError.Show(this, null, message, exception);
        }

        public void NavigateToTutorial(int index)
        {
            listViewTutorials.Items[index].Selected = true;
        }

        public Icon DisplayIcon
        {
            get 
            {
                return this.Icon;
            }
        }

        public int LCID 
        {
            get
            {
                return FormOptions.LCID;
            }
        }

        #endregion

        #region Methods

        private void ApplyConfiguration()
        {
            if (FormOptions.ConnectToDocumentation)
            {
                webBrowserTutorialContent.Visible = true;
                panelShowTutorialLink.Visible = false;
            }
            else
            {
                webBrowserTutorialContent.Visible = false;
                panelShowTutorialLink.Visible = true;
            }

            foreach (ListViewItem item in listViewTutorials.Items)
            {
                ITutorial example = item.Tag as ITutorial;
                item.Text = example.Caption;
            }

            if (listViewTutorials.SelectedItems.Count > 0)
            {
                ITutorial selectedTutorial = listViewTutorials.SelectedItems[0].Tag as ITutorial;

                labelTutorialDescription.Text = selectedTutorial.Description;

                if (FormOptions.ConnectToDocumentation)
                {
                    webBrowserTutorialContent.Navigate(new Uri(selectedTutorial.Uri));
                }
                else
                {
                    linkLabelTutorialContent.Tag = selectedTutorial.Uri;
                    linkLabelTutorialContent.Text = selectedTutorial.Uri;
                }
            }

            Translator.TranslateControls(this, "FormBase.txt");
        }

        protected internal void LoadTutorial(ITutorial tutorial)
        {
            tutorial.Connect(this);
            ListViewItem viewItem = listViewTutorials.Items.Add(tutorial.Caption);
            viewItem.ImageIndex = 0;
            viewItem.Tag = tutorial;
            _listTutorials.Add(tutorial);
            if (null != tutorial.Panel)
            { 
                tutorial.Panel.Visible = false;
                panelTutorialArea.Controls.Add(tutorial.Panel);
                tutorial.Panel.Dock = DockStyle.Fill;
            }
        }

        #endregion

        #region UI Trigger

        private void buttonRunTutorial_Click(object sender, EventArgs e)
        {
            try
            {
                buttonRunTutorial.Enabled = false;
                if (listViewTutorials.SelectedItems.Count > 0)
                {
                    ITutorial selectedTutorial = listViewTutorials.SelectedItems[0].Tag as ITutorial;
                    selectedTutorial.Run();
                }
            }
            catch (Exception exception)
            {
                FormError.Show(this, null, exception.Message, exception);
            }
            finally
            {
                buttonRunTutorial.Enabled = true;
            }
        }

        private void listViewTutorials_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if ((panelTutorialArea.Controls.Count == 0) || (listViewTutorials.SelectedItems.Count == 0))
                    return;
                
                foreach (Control item in panelTutorialArea.Controls)
                    item.Visible = false;
                ITutorial selectedTutorial = listViewTutorials.SelectedItems[0].Tag as ITutorial;

                if (null != selectedTutorial.Panel)
                    selectedTutorial.Panel.Visible = true;
                else
                    buttonRunTutorial.Visible = true;

                labelTutorialDescription.Text = selectedTutorial.Description;
                if (FormOptions.ConnectToDocumentation)
                {
                    webBrowserTutorialContent.Navigate(new Uri((listViewTutorials.SelectedItems[0].Tag as ITutorial).Uri));                    
                }
                else
                {
                    linkLabelTutorialContent.Tag = selectedTutorial.Uri;
                    linkLabelTutorialContent.Text = selectedTutorial.Uri;
                }
            }
            catch(Exception exception)
            {
                FormError.Show(this, null, exception.Message, exception);               
            }
        }
         
        private void buttonOptions_Click(object sender, EventArgs e)
        {
            try
            {
                FormOptions dialog = new FormOptions();
                if (DialogResult.OK == dialog.ShowDialog(this))
                    ApplyConfiguration();
            }
            catch(Exception exception)
            {
                FormError.Show(this, null, exception.Message, exception);  
            }
        }

        private void linkLabelMultiLanguage_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                LinkLabel label = sender as LinkLabel;
                string link = label.Tag as string;
                string[] array = link.Split(new string[] { "#" }, StringSplitOptions.RemoveEmptyEntries);
                string root = "http://netoffice.codeplex.com";
                if (FormOptions.LCID == 1033)
                    link = root + array[0];
                else
                    link = root + array[1];
                System.Diagnostics.Process.Start(link);
            }
            catch (Exception exception)
            {
                FormError.Show(this, null, exception.Message, exception);
            }
        }

        private void linkLabelTutorialContent_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                LinkLabel label = sender as LinkLabel;
                string link = label.Tag as string;
                System.Diagnostics.Process.Start(link);
            }
            catch (Exception exception)
            {
                FormError.Show(this, null, exception.Message, exception);
            }
        }

        private void linkLabelDiscussion_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                LinkLabel label = sender as LinkLabel;
                string link = label.Tag as string;
                System.Diagnostics.Process.Start(link);
            }
            catch (Exception exception)
            {
                FormError.Show(this, null, exception.Message, exception);
            }
        }

        private void FormBase_FormClosed(object sender, FormClosedEventArgs e)
        {
            foreach (ITutorial item in _listTutorials)
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

            FormOptions.SaveConfigurationToXMLFile();
        }

        private void FormBase_Resize(object sender, EventArgs e)
        {
            try
            {
                buttonRunTutorial.Left = (buttonRunTutorial.Parent.Width) / 2 - (buttonRunTutorial.Width / 2);
                buttonRunTutorial.Top = (buttonRunTutorial.Parent.Height) / 2 - (buttonRunTutorial.Height / 2);
            }
            catch (Exception exception)
            {
                FormError.Show(this, null, exception.Message, exception);
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            FormBase_Resize(this, new EventArgs());
        }

        #endregion  
    }
}
