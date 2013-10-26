using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NOTools.CSharpTextEditorPFCreator
{
    public partial class FormCreator : Form
    {
        #region Ctor

        public FormCreator()
        {
            InitializeComponent();
            SetDefaultResultFolder();
        }

        #endregion

        #region Methods

        private void SetDefaultResultFolder()
        {
            string folder = System.IO.Path.Combine(Application.StartupPath, "PF-Files");
            textBoxResultFolder.Text = folder;
        }

        private void OpenFolder()
        {
            if (System.IO.Directory.Exists(textBoxResultFolder.Text))
                System.Diagnostics.Process.Start(textBoxResultFolder.Text);
        }

        private bool CheckValidSettings()
        {
            if (String.IsNullOrWhiteSpace(textBoxResultFolder.Text))
                return false;

            if (listViewAssemblies.Items.Count == 0)
                return false;

            return true;
        }

        private bool PerformFolderCheck()
        {
            if (!System.IO.Directory.Exists(textBoxResultFolder.Text))
            {
                DialogResult dr = MessageBox.Show(this, "Result folder doesnt exists. Create anyway?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if(dr != System.Windows.Forms.DialogResult.Yes)
                    return false;
            }

            if ((System.IO.Directory.Exists(textBoxResultFolder.Text)) && System.IO.Directory.EnumerateFiles(textBoxResultFolder.Text).Count() > 0)
            {
                DialogResult dr = MessageBox.Show(this, "The result folder is not empty. Continue?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr != System.Windows.Forms.DialogResult.Yes)
                    return false;
            }

            return true;
        }

        private string[] CreateAssemblyArray()
        {
            List<string> list = new List<string>();
            foreach (ListViewItem item in listViewAssemblies.Items)
                list.Add(item.SubItems[1].Text);
            return list.ToArray();
        }

        private bool ListViewContainsItem(string text)
        {
            foreach (ListViewItem item in listViewAssemblies.Items)
            {
                if (item.Text.Equals(text, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            }
            return false;
        }

        #endregion

        #region Trigger
        
        private void buttonStart_Click(object sender, EventArgs e)
        {
            try
            {
                if (!CheckValidSettings())
                {
                    MessageBox.Show("Invalid Settings.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!PerformFolderCheck())
                    return;

                this.Cursor = Cursors.WaitCursor;
                DateTime start = DateTime.Now;
                PersistenceFileCreator.CreatePersistenceFiles(textBoxResultFolder.Text, CreateAssemblyArray(), checkBoxCreateCompressedCopies.Checked);
                TimeSpan timeElapsed = DateTime.Now.Subtract(start);

                MessageBox.Show(this, "Done! Time elapsed: " + timeElapsed.ToString(), this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                OpenFolder();

            }
            catch (Exception exception)
            {
                MessageBox.Show(this, "Error: " + exception.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;
            dialog.Filter = "Libraries (*.dll)|*.dll|All Files(*.*)|*.*";
            if (DialogResult.OK == dialog.ShowDialog(this))
            {
                foreach (string file in dialog.FileNames)
                {
                    string fileName = System.IO.Path.GetFileName(file);
                    if (!ListViewContainsItem(fileName))
                    { 
                        ListViewItem item = listViewAssemblies.Items.Add(fileName);
                        item.SubItems.Add(file);
                    }
                }               
            }
        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            if (0 == listViewAssemblies.SelectedItems.Count)
                return;

            ListViewItem item = listViewAssemblies.SelectedItems[0];
            item.Remove();
        }

        private void buttonChooseFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (DialogResult.OK == dialog.ShowDialog(this))
                textBoxResultFolder.Text = dialog.SelectedPath;
        }

        private void listViewAssemblies_SelectedIndexChanged(object sender, EventArgs e)
        {
            buttonDelete.Enabled = listViewAssemblies.SelectedItems.Count > 0;
        }

        #endregion
    }
}
