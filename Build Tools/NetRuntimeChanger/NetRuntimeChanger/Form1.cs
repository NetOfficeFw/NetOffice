using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace NetRuntimeChanger
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            comboBoxVersion.SelectedIndex = 0;
            comboBoxLanguage.SelectedIndex = 0;
        }

        private void buttonStart_Click(object sender, EventArgs e)
        {
            string projectExtension = "";
            
            if(comboBoxLanguage.SelectedIndex ==0)
                projectExtension = "*.csproj";
            else
                projectExtension = "*.vbproj";

            textBoxLog.Text = "";
            string netVersion = GetSelectedNetVersion();

            if (!Directory.Exists(textBoxFolder.Text))
            {
                MessageBox.Show("Directory not exists", "Doooh", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //change projects
            string[] files = Directory.GetFiles(textBoxFolder.Text, projectExtension, SearchOption.AllDirectories);
            foreach (string file in files)
            {
                textBoxLog.AppendText("Change : " + System.IO.Path.GetFileName(file) + "\r\n");
                string fileContent = File.ReadAllText(file);
                
                ChangeNetVersionEntry(ref fileContent, netVersion);
                ChangeToolsVersionEntry(ref fileContent, netVersion);
                ChangeKeyFile(ref fileContent, System.IO.Path.GetDirectoryName(file), System.IO.Path.GetFileNameWithoutExtension (file), netVersion);
               
                File.Delete(file);
                File.WriteAllText(file, fileContent, Encoding.UTF8);
            }

            //change solution
            string[] slnFiles = Directory.GetFiles(textBoxFolder.Text, "*.sln", SearchOption.AllDirectories);
            foreach (string slnFile in slnFiles)
            {
                textBoxLog.AppendText("Change : " + System.IO.Path.GetFileName(slnFile) + "\r\n");
                string fileContent = File.ReadAllText(slnFile);

                ChangeFormatVersion(ref fileContent, netVersion);
                ChangeToolCode(ref fileContent, netVersion);

                File.Delete(slnFile);
                File.WriteAllText(slnFile, fileContent, Encoding.UTF8);
            }

        }

        private void ChangeKeyFile(ref string fileContent, string currentFolder, string name, string netVersion)
        {
            if (!checkBoxChangeKeyFiles.Checked)
                return;

            int position1 = fileContent.IndexOf("<AssemblyOriginatorKeyFile>");
            if (position1 < 0)
            {
                textBoxLog.AppendText("\t\tKeyFile Entry not found:" + name + "\r\n");
                return;
            }

            int position2 = fileContent.IndexOf("</AssemblyOriginatorKeyFile>", position1);
            string keyFile = fileContent.Substring(position1 + "<AssemblyOriginatorKeyFile>".Length, position2 - (position1 + "<AssemblyOriginatorKeyFile>".Length));
            string[] arr = keyFile.Split(new string[] {"_"},StringSplitOptions.RemoveEmptyEntries);
            string newKeyFile = arr[0] + "_v" + netVersion + ".snk";
            string newLine = "<AssemblyOriginatorKeyFile>" + newKeyFile + "</AssemblyOriginatorKeyFile>";

            string net1 = "<AssemblyOriginatorKeyFile>" + keyFile + "</AssemblyOriginatorKeyFile>";
            fileContent = fileContent.Replace(net1, newLine);

            string sourceFile = System.IO.Path.Combine(textBoxKeyFilesRootFolder.Text, netVersion);
            sourceFile = System.IO.Path.Combine(sourceFile, newKeyFile);

            string destFile = System.IO.Path.Combine(currentFolder, newKeyFile);
            if(!System.IO.File.Exists(destFile))
                System.IO.File.Copy(sourceFile, destFile);
        }

        private void ChangeToolCode(ref string fileContent, string netVersion)
        {
            string replacedString = "";
            if (netVersion == "4.0")
            {
                if (comboBoxLanguage.SelectedIndex == 0)
                    replacedString = "# Visual C# Express 2010";
                else
                    replacedString = "# Visual Basic Express 2010";
            }
            else
                replacedString = "# Visual Studio 2008";

            string tools1 = "# Visual C# Express 2010";
            int position = fileContent.IndexOf(tools1);
            if (position > -1)
            {
                fileContent = fileContent.Replace(tools1, replacedString);
                return;
            }

            string tools2 = "# Visual Studio 2008";
            position = fileContent.IndexOf(tools2);
            if (position > -1)
            {
                fileContent = fileContent.Replace(tools2, replacedString);
                return;
            }
        }

        private void ChangeFormatVersion(ref string fileContent, string netVersion)
        {
            string replacedString = "";
            if (netVersion.StartsWith("4"))
                replacedString = "Format Version 11.00";
            else
                replacedString = "Format Version 10.00";

            string tools1 = "Format Version 10.00";
            int position = fileContent.IndexOf(tools1);
            if (position > -1)
            {
                fileContent = fileContent.Replace(tools1, replacedString);
                return;
            }

            string tools2 = "Format Version 11.00";
            position = fileContent.IndexOf(tools2);
            if (position > -1)
            {
                fileContent = fileContent.Replace(tools2, replacedString);
                return;
            }
        }

        private void ChangeToolsVersionEntry(ref string fileContent, string netVersion)
        {
            string replacedString = "";
            if (netVersion.StartsWith("4"))
                replacedString = "ToolsVersion=\"4.0\"";
            else
                replacedString = "ToolsVersion=\"3.5\"";

            string tools1 = "ToolsVersion=\"3.5\"";
            int position = fileContent.IndexOf(tools1);
            if (position > -1)
            {
                fileContent = fileContent.Replace(tools1, replacedString);
                return;
            }

            string tools2 = "ToolsVersion=\"4.0\"";
            position = fileContent.IndexOf(tools2);
            if (position > -1)
            {
                fileContent = fileContent.Replace(tools2, replacedString);
                return;
            }
        }

        private void ChangeNetVersionEntry(ref string fileContent, string netVersion)
        {
            string net1 = "<TargetFrameworkVersion>v2.0</TargetFrameworkVersion>";
            int position = fileContent.IndexOf(net1);
            if (position > -1)
            {
                fileContent = fileContent.Replace(net1, "<TargetFrameworkVersion>v" + netVersion + "</TargetFrameworkVersion>");
                return;
            }

            string net2 = "<TargetFrameworkVersion>v3.0</TargetFrameworkVersion>";
            position = fileContent.IndexOf(net2);
            if (position > -1)
            {
                fileContent = fileContent.Replace(net2, "<TargetFrameworkVersion>v" + netVersion + "</TargetFrameworkVersion>");
                return;
            }

            string net3 = "<TargetFrameworkVersion>v3.5</TargetFrameworkVersion>";
            position = fileContent.IndexOf(net3);
            if (position > -1)
            {
                fileContent = fileContent.Replace(net3, "<TargetFrameworkVersion>v" + netVersion + "</TargetFrameworkVersion>");
                return;
            }

            string net4 = "<TargetFrameworkVersion>v4.0</TargetFrameworkVersion>";
            position = fileContent.IndexOf(net4);
            if (position > -1)
            {
                fileContent = fileContent.Replace(net4, "<TargetFrameworkVersion>v" + netVersion + "</TargetFrameworkVersion>");
                return;
            }


            string net5 = "<TargetFrameworkVersion>v4.5</TargetFrameworkVersion>";
            position = fileContent.IndexOf(net5);
            if (position > -1)
            {
                fileContent = fileContent.Replace(net4, "<TargetFrameworkVersion>v" + netVersion + "</TargetFrameworkVersion>");
                return;
            }
        }

        private string GetSelectedNetVersion()
        {
            switch (comboBoxVersion.SelectedIndex)
            { 
                case 0:
                    return "2.0";
                case 1:
                    return "3.0";
                case 2:
                    return "3.5";
                case 3:
                    return "4.0";
                case 4:
                    return "4.5";
            }
            throw(new Exception("Version not selected."));
        }

        private void buttonChooseFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fdg = new FolderBrowserDialog();
            if (DialogResult.OK == fdg.ShowDialog(this))
                textBoxFolder.Text = fdg.SelectedPath;

        }
    }
}
