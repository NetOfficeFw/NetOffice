using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NOTools.CSharpTextEditor;
using NOTools.InMemoryCompiler;

namespace NOTools.CSharpTextEditorPFCreator
{
    public partial class FormTest : Form
    {
        #region Ctor

        public FormTest()
        {
            InitializeComponent();
            SetDefaultResultFolder();
             
        }

        #endregion

        #region Methods

        private void SetDefaultResultFolder()
        {
            string folder = System.IO.Path.Combine(Application.StartupPath, "XPF");
            textBoxResultFolder.Text = folder;
        }

        private bool CheckValidSettings()
        {
            if (String.IsNullOrWhiteSpace(textBoxResultFolder.Text))
                return false;

            if (!System.IO.Directory.Exists(textBoxResultFolder.Text))
                return false;

            if (listViewAssemblies.Items.Count == 0)
                return false;

            return true;
        }

        private string[] CreateAssemblyNameArray()
        {
            List<string> list = new List<string>();
            foreach (ListViewItem item in listViewAssemblies.Items)
            {
                string name = System.IO.Path.GetFileNameWithoutExtension(item.Text);
                list.Add(name);
            }
            return list.ToArray();
        }

        private string[] CreateAssemblyPathArray()
        {
            List<string> list = new List<string>();
            foreach (ListViewItem item in listViewAssemblies.Items)
            {
                list.Add(item.SubItems[1].Text);
            }
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

        private void codeEditorControl1_CompileRequest(CodeEditorControl sender, CompileRequestEventArgs args)
        {
            codeEditorControl1.ErrorPanelSettings.Header = "Compile...";
            DynamicAssembly assembly = new DynamicAssembly("TestAssembly", codeEditorControl1.References.ToStringPathArray());
            assembly.CustomClasses.AddNew(codeEditorControl1.Text);
            CompileResult result = CSharpCompiler.CompileDynamicAssembly(assembly);
            codeEditorControl1.ShowErrors(result.Errors, "Sucseed");
        }
         
        private void buttonStart_Click(object sender, EventArgs e)
        {
            try
            {
                if (!CheckValidSettings())
                {
                    MessageBox.Show("Invalid Settings.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                DateTime start = DateTime.Now;
                codeEditorControl1.PersistencePath = textBoxResultFolder.Text;

                if (checkBoxAsync.Checked)
                    codeEditorControl1.AddReferencesFromFile(CreateAssemblyNameArray(), CreateAssemblyPathArray(), true, true);
                else
                {
                    codeEditorControl1.AddReferencesFromFile(CreateAssemblyNameArray(), CreateAssemblyPathArray(), true, false);
                    TimeSpan timeElapsed = DateTime.Now.Subtract(start);
                    MessageBox.Show(this, "Done! Time elapsed: " + timeElapsed.ToString(), this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, "Error: " + exception.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            if (0 == listViewAssemblies.SelectedItems.Count)
                return;

            ListViewItem item = listViewAssemblies.SelectedItems[0];
            item.Remove();
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

        private void buttonChooseFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (DialogResult.OK == dialog.ShowDialog(this))
                textBoxResultFolder.Text = dialog.SelectedPath;
        }

        private void listViewAssemblies_SelectedIndexChanged(object sender, EventArgs e)
        {
            buttonDelete.Enabled = listViewAssemblies.Items.Count > 0;
        }

        #endregion
    }
}
