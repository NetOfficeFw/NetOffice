using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ExampleBase
{
    public partial class ExampleForm : Form, IHost
    {
        #region Fields

        private string _rootDirectory = OptionsForm.DefaultRootDirectory;

        #endregion

        #region Ctor

        public ExampleForm()
        {
            InitializeComponent();
            ExamplesView.AutoGenerateColumns = false;
            DataSource = new SortableBindingList<ExampleViewItem>();
            ExamplesView.DataSource = DataSource;
        }

        #endregion

        #region Properties

        private SortableBindingList<ExampleViewItem> DataSource { get; set; }
        
        private IExample SelectedExample
        {
            get
            {
                if (ExamplesView.SelectedRows.Count > 0)
                    return (ExamplesView.SelectedRows[0].DataBoundItem as ExampleViewItem).Item;
                else
                    return null;
            }
        }

        #endregion

        #region IHost

        public DialogResult ShowQuestion(string message)
        {
            return MessageBox.Show(this, message, "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        }

        public void ShowMessage(string message)
        {
            MessageBox.Show(this, message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void ShowFinishDialog(string message, string fullDocumentPath)
        {
            try
            {
                FinishForm dialog = new FinishForm(message, fullDocumentPath);
                dialog.ShowDialog(this);
            }
            catch(Exception exception)
            {
                ErrorForm.Show(this, exception);           
            }
        }

        public void ShowErrorDialog(string message, Exception exception)
        {
            ErrorForm.Show(this, exception);  
        }

        public Icon DisplayIcon
        {
            get 
            {
                return this.Icon;
            }
        }

        public string RootDirectory
        {
            get
            {
                return _rootDirectory;
            }
        }

        #endregion

        #region Methods

        protected internal void LoadExample(IExample example)
        {
            example.Connect(this);
            DataSource.Add(new ExampleViewItem(example, FormImageList.Images[0]));
        }

        private Label ShowWait()
        {
            Label label = new Label();
            label.TextAlign = ContentAlignment.MiddleCenter;
            label.Text = "Processing...";
            label.Font = HeaderLabel.Font;
            label.Location = ExamplesView.Location;
            label.Size = ExamplesView.Size;
            Controls.Add(label);
            label.BringToFront();
            return label;
        }

        private void HideWait(Label label)
        {
            if (null != label)
            {
                Controls.Remove(label);
                label.Dispose();
            }
        }

        #endregion

        #region Trigger

        private void ExamplesView_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            var example = SelectedExample;
            if (null == example)
                return;

            Label label = null;
            try
            {
                Cursor = Cursors.WaitCursor;
                Enabled = false;

                if (null != example.Panel)
                    AreaForm.ShowForm(this, example);
                else
                {
                    label = ShowWait();
                    example.RunExample();
                }
            }
            catch (Exception exception)
            {
                ErrorForm.Show(this, exception);
            }
            finally
            {
                HideWait(label);
                Enabled = true;
                Cursor = Cursors.Default;
            }
        }
       
        private void OptionButton_Click(object sender, EventArgs e)
        {
            try
            {
                OptionsForm dialog = new OptionsForm(_rootDirectory);
                if (DialogResult.OK == dialog.ShowDialog(this))
                {
                    _rootDirectory = dialog.RootDirectory;
                }
            }
            catch (Exception exception)
            {
                ErrorForm.Show(this, exception);
            }
        }

        private void TopicLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
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

        #endregion
    }
}
