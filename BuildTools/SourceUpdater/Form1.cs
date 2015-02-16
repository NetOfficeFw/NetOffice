using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace NOBuildTools.SourceUpdater
{
    /// <summary>
    /// Main form in the application
    /// </summary>
    public partial class Form1 : Form
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public Form1()
        {
            InitializeComponent();
        }

        #endregion

        #region Methods

        private bool GetIsInSvnFolder(string file)
        {
            if (file.IndexOf(".svn",StringComparison.InvariantCultureIgnoreCase) > -1)
                return true;
            else
                return false;

        }

        #endregion

        #region Trigger

        private void buttonStart_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(textBoxSource.Text))
            {
                MessageBox.Show("SourceDir not found");
                return;
            }

            if (!Directory.Exists(textBoxDest.Text))
            {
                MessageBox.Show("DestDir not found");
                return;
            }

            int i = 0;
            string[] codeFiles = Directory.GetFiles(textBoxSource.Text, "*.*", SearchOption.AllDirectories);
            foreach (string item in codeFiles)
            {
                bool isInSvnFolder = GetIsInSvnFolder(item);
                if (false == isInSvnFolder)
                {
                    textBoxLog.Text = item;
                    string newFilePath = GetNewFilePath(textBoxSource.Text, textBoxDest.Text, item);
                    if (File.Exists(newFilePath))
                        File.Delete(newFilePath);

                    if (!Directory.Exists(Path.GetDirectoryName(newFilePath)))
                        Directory.CreateDirectory(Path.GetDirectoryName(newFilePath));

                    File.Copy(item, newFilePath);
                    i++;
                }
            }

            textBoxLog.Text = "Finish " + i.ToString() + " Files.";
        }

        private string GetNewFilePath(string sourceDir, string DestDir, string file)
        {
            file = file.Substring(sourceDir.Length);
            return DestDir + file;
        }

        private void buttonChooseSource_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (DialogResult.OK == fbd.ShowDialog(this))
                textBoxSource.Text = fbd.SelectedPath;  
        }

        private void buttonChooseDest_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (DialogResult.OK == fbd.ShowDialog(this))
                textBoxDest.Text = fbd.SelectedPath;
        }

        #endregion

    }
}
