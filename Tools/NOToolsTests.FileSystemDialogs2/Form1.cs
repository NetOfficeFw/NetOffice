using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NOToolsTests.FileSystemDialogs2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
           
        }

        private void openFilePanel1_FileDoubleClick(object sender, NOTools.FileSystemDialogs.FileDoubleClickEventArgs args)
        {
            MessageBox.Show(this, String.Format("Double click on: {1} {0}", args.File, Environment.NewLine));
        }
        
        private void openFilePanel1_SelectionChanged(object sender, NOTools.FileSystemDialogs.SelectionChangedEventArgs args)
        {
            ButtonSelectFile.Enabled = args.Files.Length > 0;
        }

        private void ButtonSelectFile_Click(object sender, EventArgs e)
        {
            string clickedFiles = string.Empty;
            foreach (var item in openFilePanel1.Misc.SelectedFiles)
                clickedFiles += item + Environment.NewLine;
            MessageBox.Show(this, String.Format("You have selected: {1} {0}", clickedFiles, Environment.NewLine));
        }
    }
}
