using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NOTools.FileSystemDialogs;

namespace NOToolsTests.FileSystemDialogs3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void buttonOpen_Click(object sender, EventArgs e)
        {
            var dr = OpenFilePanelDialog.Show(this, "All(*.*)|*.*");
            if (dr.Result != System.Windows.Forms.DialogResult.Cancel)
            {
                string clickedFiles = string.Empty;
                foreach (var item in dr.SelectedFiles)
                    clickedFiles += item + Environment.NewLine;
                MessageBox.Show(this, String.Format("You have selected: {1} {0}", clickedFiles, Environment.NewLine));
            }
        }
    }
}
