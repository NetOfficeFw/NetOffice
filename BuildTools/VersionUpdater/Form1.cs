using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace NOBuildTools.VersionUpdater
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            comboBoxFromNetVersion.SelectedIndex = 0;
            comboBoxToNetVersion.SelectedIndex = 0;             
        }

        private string GetSelectedFromNetVersion()
        {
            switch (comboBoxFromNetVersion.SelectedIndex)
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
            throw (new Exception("From .NET Version not selected."));
        }

        private string GetSelectedToNetVersion()
        {
            switch (comboBoxToNetVersion.SelectedIndex)
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
            throw (new Exception("To .NET Version not selected."));
        }

        private bool CheckPrequsits()
        {
            if (!Directory.Exists(textBoxFolder.Text))
            {
                MessageBox.Show("Directory not exists", "Doooh", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (GetSelectedFromNetVersion() == GetSelectedToNetVersion())
            {
                MessageBox.Show("Same .NET versions.", "Doooh", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (!Directory.Exists(textBoxNewFolder.Text))
            {
                Directory.CreateDirectory(textBoxNewFolder.Text);
            }
            else
            {
                if (Directory.GetFiles(textBoxNewFolder.Text, "*.*").Length > 0 ||
                    Directory.GetDirectories(textBoxNewFolder.Text).Length > 0)
                {
                    DialogResult dr = MessageBox.Show(this, String.Format("The is folder{0}{1}{0}ist not empty and want be deleted.{0}Continue anyway?", Environment.NewLine, textBoxNewFolder.Text), "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                    if (dr == System.Windows.Forms.DialogResult.No)
                        return false;
                }
            }
            return true;
        }

        private void buttonStart_Click(object sender, EventArgs e)
        {
            if (!CheckPrequsits())
                return;

            textBoxLog.Clear();
            string netFromVersion = GetSelectedFromNetVersion();
            string netToVersion = GetSelectedToNetVersion();

            //change projects
            string[] files = Directory.GetFiles(textBoxFolder.Text, "*.*", SearchOption.AllDirectories);
            foreach (string file in files)
            {
                if (file.EndsWith("vbproj") || file.EndsWith("csproj"))
                {
                    textBoxLog.AppendText("Change : " + System.IO.Path.GetFileName(file) + "\r\n");
                    string fileContent = File.ReadAllText(file);

                    ChangeNetVersionEntryInProjectFile(ref fileContent, netToVersion);
                    ChangeToolsVersionEntryInProjectFile(ref fileContent);
                    ChangeKeyFileInProjectFile(ref fileContent, System.IO.Path.GetDirectoryName(file), System.IO.Path.GetFileNameWithoutExtension(file), netToVersion);

                    File.Delete(file);
                    File.WriteAllText(file, fileContent, Encoding.UTF8);
                }
            }

            //change solution
            string[] slnFiles = Directory.GetFiles(textBoxFolder.Text, "*.sln", SearchOption.AllDirectories);
            foreach (string slnFile in slnFiles)
            {
                textBoxLog.AppendText("Change : " + System.IO.Path.GetFileName(slnFile) + "\r\n");

                string fileContent = File.ReadAllText(slnFile);
                ChangeFormatVersionInSolutionFile(ref fileContent, netToVersion);
                if(fileContent.IndexOf(".csproj") > -1)
                    ChangeToolCodeInCSharpSolutionFile(ref fileContent, netToVersion);
                else if (fileContent.IndexOf(".vbproj") > -1)
                    ChangeToolCodeInVBSolutionFile(ref fileContent, netToVersion);
                else
                {
                    throw new IndexOutOfRangeException();
                }

                File.Delete(slnFile);
                File.WriteAllText(slnFile, fileContent, Encoding.UTF8);
            }

            //change net marker
            if (!checkBoxChangeNetMarker.Checked)
                return;
            string[] allFiles = Directory.GetFiles(textBoxFolder.Text, "*.*", SearchOption.AllDirectories);
            foreach (string file in allFiles)
            {
                if (file.EndsWith(".cs") || file.EndsWith(".csproj"))
                {
                    string fileContent = File.ReadAllText(file);
                    string csMarkerSource = "CS" + netFromVersion.Replace(".", "");
                    string csMarkerTarget = "CS" + netToVersion.Replace(".", "");
                    if (fileContent.IndexOf(csMarkerSource) > -1)
                    {
                        fileContent = fileContent.Replace(csMarkerSource, csMarkerTarget);
                        File.Delete(file);
                        File.WriteAllText(file, fileContent, Encoding.UTF8);
                    }
                }
                else if (file.EndsWith(".vb") || file.EndsWith(".vbproj"))
                {
                    string fileContent = File.ReadAllText(file);
                    string vbMarkerSource = "VB" + netFromVersion.Replace(".", "");
                    string vbMarkerTarget = "VB" + netToVersion.Replace(".", "");
                    if (fileContent.IndexOf(vbMarkerSource) > -1)
                    {
                        fileContent = fileContent.Replace(vbMarkerSource, vbMarkerTarget);
                        File.Delete(file);
                        File.WriteAllText(file, fileContent, Encoding.UTF8);
                    }
                }                
            }
        }

        private void ChangeKeyFileInProjectFile(ref string fileContent, string currentFolder, string name, string netVersion)
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
        
        private void ChangeToolCodeInCSharpSolutionFile(ref string fileContent, string toNetVersion)
        {
            string replaceToolsString = "# Visual C# Express 2010";

            string tools2008 = "# Visual Studio 2008";
            int position = fileContent.IndexOf(tools2008);
            if (position > -1)
            {
                fileContent = fileContent.Replace(tools2008, replaceToolsString);
                return;
            }

            string tools2010 = "# Visual Studio 2010";
            position = fileContent.IndexOf(tools2010);
            if (position > -1)
            {
                fileContent = fileContent.Replace(tools2010, replaceToolsString);
                return;
            }
        }

        private void ChangeToolCodeInVBSolutionFile(ref string fileContent, string toNetVersion)
        {
            string replaceToolsString = "# Visual Basic Express 2010";

            string tools2008 = "# Visual Studio 2008";
            int position = fileContent.IndexOf(tools2008);
            if (position > -1)
            {
                fileContent = fileContent.Replace(tools2008, replaceToolsString);
                return;
            }

            string tools2010 = "# Visual Studio 2010";
            position = fileContent.IndexOf(tools2010);
            if (position > -1)
            {
                fileContent = fileContent.Replace(tools2010, replaceToolsString);
                return;
            }
        }


        //private void ChangeToolCodeInSolutionFile(ref string fileContent, string netVersion)
        //{
        //    string replacedString = "";
        //    if (comboBoxLanguage.SelectedIndex == 0)
        //        replacedString = "# Visual C# Express 2010";
        //    else
        //        replacedString = "# Visual Basic Express 2010";
                
        //    string tools1 = "# Visual C# Express 2010";
        //    int position = fileContent.IndexOf(tools1);
        //    if (position > -1)
        //    {
        //        fileContent = fileContent.Replace(tools1, replacedString);
        //        return;
        //    }

        //    string tools2 = "# Visual Studio 2008";
        //    position = fileContent.IndexOf(tools2);
        //    if (position > -1)
        //    {
        //        fileContent = fileContent.Replace(tools2, replacedString);
        //        return;
        //    }

        //    string tools3 = "# Visual Studio 2010";
        //    position = fileContent.IndexOf(tools3);
        //    if (position > -1)
        //    {
        //        fileContent = fileContent.Replace(tools3, replacedString);
        //        return;
        //    }

        //    string tools4 = "# Visual Basic Express 2010";
        //    position = fileContent.IndexOf(tools4);
        //    if (position > -1)
        //    {
        //        fileContent = fileContent.Replace(tools4, replacedString);
        //        return;
        //    }
        //}

        private void ChangeFormatVersionInSolutionFile(ref string fileContent, string netVersion)
        {
            string replacedString = "";
            replacedString = "Format Version 11.00";

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

        private void ChangeToolsVersionEntryInProjectFile(ref string fileContent)//, string netVersion)
        {
            string replacedString = "";
            replacedString = "ToolsVersion=\"4.0\"";

            string searchTools1 = "ToolsVersion=\"3.5\"";
            int position = fileContent.IndexOf(searchTools1);
            if (position > -1)
            {
                fileContent = fileContent.Replace(searchTools1, replacedString);
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

        private void ChangeNetVersionEntryInProjectFile(ref string fileContent, string netVersion)
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
   
        private void buttonChooseFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fdg = new FolderBrowserDialog();
            if (DialogResult.OK == fdg.ShowDialog(this))
                textBoxFolder.Text = fdg.SelectedPath;

        }

        private void buttonChooseNewFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fdg = new FolderBrowserDialog();
            if (DialogResult.OK == fdg.ShowDialog(this))
                textBoxNewFolder.Text = fdg.SelectedPath;
        }
    }
}
