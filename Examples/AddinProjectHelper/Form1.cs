using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace AddinProjectHelper
{
    public partial class Form1 : Form
    {
        #region Ctor

        public Form1()
        {
            InitializeComponent();
            buttonInject.Enabled = !String.IsNullOrWhiteSpace(textBoxPath.Text);
        }

        #endregion

        #region Properties

        private string SelectedPath
        {
            get
            {
                return textBoxPath.Text;
            }
            set
            {
                textBoxPath.Text = value;
                buttonInject.Enabled = !String.IsNullOrWhiteSpace(textBoxPath.Text);
            }
        }

        #endregion

        #region Methods

        #endregion

        #region Trigger

        private void buttonExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void folderButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog(this) == DialogResult.OK)
                SelectedPath = dlg.SelectedPath;
        }

        private void buttonInject_Click(object sender, EventArgs e)
        {
            try
            {
                string path = Application.StartupPath;
                var injector = new Injector();
                injector.Inject(path, "Excel", SelectedPath);
                injector.Inject(path, "Word", SelectedPath);
                injector.Inject(path, "Outlook", SelectedPath);
                injector.Inject(path, "PowerPoint", SelectedPath);
                injector.Inject(path, "Access", SelectedPath);
                MessageBox.Show(this, "Done!", Text);
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.ToString(), "Error");
            }
        }

        #endregion
    }
}