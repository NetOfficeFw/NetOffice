using System;
using System.Windows.Forms;

namespace ExampleBase
{
    /// <summary>
    /// Application config options dialog
    /// </summary>
    partial class OptionsForm : Form
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="rootDirectory">current output directory</param>
        public OptionsForm(string rootDirectory)
        {
            InitializeComponent();
            if (Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) != rootDirectory)
                radioButtonApplicationFolder.Checked = true;
        }
        
        /// <summary>
        /// Current output directory for created office files
        /// </summary>
        public string RootDirectory
        {
            get
            {
                return radioButtonCommonFolder.Checked ? 
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) : Application.StartupPath;
            }
        }

        /// <summary>
        /// Default output directory for created office files
        /// </summary>
        public static string DefaultRootDirectory
        {
            get 
            {
                return Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            }
        }

        #endregion

        #region Trigger

        private void buttonDone_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        #endregion
    }
}