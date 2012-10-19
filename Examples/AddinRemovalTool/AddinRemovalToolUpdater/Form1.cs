using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace AddinRemovalToolUpdater
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void buttonDoAction_Click(object sender, EventArgs e)
        {
            try
            {
                textBox3.Clear();
                string assemblyFileLocation = textBox1.Text;
                string releaseFolderLocation = textBox2.Text;
                string assemblyFileName = Path.GetFileName(assemblyFileLocation);

                if (!File.Exists(assemblyFileLocation))
                    throw new FileNotFoundException("Assembly file not found.", assemblyFileLocation, null);

                if (!Directory.Exists(releaseFolderLocation))
                    throw new DirectoryNotFoundException("Release folder not found.");
                
                string[] runtimes = new string[]{"NET 2.0", "NET 3.0", "NET 3.5", "NET 4.0", "NET 4.5"};
                string[] languages = new string[] { "C#", "VB" };
                string[] products = new string[] { "Access", "Excel", "Outlook", "PowerPoint", "Word", "Misc"};

                int counter = 0;
                foreach (string runtime in runtimes)
                {
                    foreach (string product in products)
                    {
                        foreach (string language in languages)
                        {
                            string targetRelativeFilePath = string.Format("{0}\\Examples\\{1}\\{2}\\{3}", runtime, product, language, assemblyFileName);
                            string targetAbsoluteFilePath = Path.Combine(releaseFolderLocation, targetRelativeFilePath);

                            if (File.Exists(targetAbsoluteFilePath))
                                textBox3.Text = "Update File " + targetRelativeFilePath + Environment.NewLine + textBox3.Text;
                            else
                                textBox3.Text = "Insert File " + targetRelativeFilePath + Environment.NewLine + textBox3.Text;

                            File.Copy(assemblyFileLocation, targetAbsoluteFilePath, true);
                            counter++;
                        }
                    }
                }
                textBox3.Text = string.Format("Proceed {0} Files{1}", counter, Environment.NewLine) + textBox3.Text;
            }
            catch (Exception exception)
            {
                MessageBox.Show("An error ocurred. " + exception.Message, "Doooh", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
