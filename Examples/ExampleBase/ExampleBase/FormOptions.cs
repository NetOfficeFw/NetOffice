using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace ExampleBase
{
    partial class FormOptions : Form
    {
        public FormOptions(int currentLCID, string rootDirectory)
        {
            InitializeComponent();

            if (1031 == currentLCID)
                radioButtonLanguage1031.Checked = true;

            if (Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) != rootDirectory)
                radioButtonApplicationFolder.Checked = true;
        }

        private void buttonDone_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        
        public int LCID
        {
            get 
            {
                return radioButtonLanguage1031.Checked ? 1031 : 1033;
            }
        }

        public string RootDirectory
        {
            get
            {
                return radioButtonCommonFolder.Checked ? 
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) : Application.StartupPath;
            }
        }

        public static int DefaultLCID
        {
            get 
            {
                return 1033;
            }
        }

        public static string DefaultRootDirectory
        {
            get 
            {
                return Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            }
        }
    }
}
