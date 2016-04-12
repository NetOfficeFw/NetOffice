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
            textBoxFolder.Text = Application.StartupPath;
            comboBoxFromNetVersion.SelectedIndex = 0;
            comboBoxToNetVersion.SelectedIndex = 0;             
        }

        #endregion

        #region Methods
        
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

            return true;
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
            if (toNetVersion == "4.5")
            {
                string replaceToolsString = "# Visual Studio Express 2012 for Windows Desktop";

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
            else
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
        }

        private void ChangeToolCodeInVBSolutionFile(ref string fileContent, string toNetVersion)
        {
            if (toNetVersion == "4.5")
            {
                string replaceToolsString = "# Visual Studio Express 2012 for Windows Desktop";

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
            else
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
        }
         
        private void ChangeToolsVersionEntryInProjectFile(ref string fileContent)
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
        }

        private void ChangeNetVersionEntryInProjectFile(ref string fileContent, string toNetVersion)
        {
            if (toNetVersion != "4.0")
            {
                fileContent = fileContent.Replace("<TargetFrameworkProfile>Client</TargetFrameworkProfile>", "");
            }


            string net1 = "<TargetFrameworkVersion>v2.0</TargetFrameworkVersion>";
            int position = fileContent.IndexOf(net1);
            if (position > -1)
            {
                fileContent = fileContent.Replace(net1, "<TargetFrameworkVersion>v" + toNetVersion + "</TargetFrameworkVersion>");
                return;
            }

            string net2 = "<TargetFrameworkVersion>v3.0</TargetFrameworkVersion>";
            position = fileContent.IndexOf(net2);
            if (position > -1)
            {
                fileContent = fileContent.Replace(net2, "<TargetFrameworkVersion>v" + toNetVersion + "</TargetFrameworkVersion>");
                return;
            }

            string net3 = "<TargetFrameworkVersion>v3.5</TargetFrameworkVersion>";
            position = fileContent.IndexOf(net3);
            if (position > -1)
            {
                fileContent = fileContent.Replace(net3, "<TargetFrameworkVersion>v" + toNetVersion + "</TargetFrameworkVersion>");
                return;
            }

            string net4 = "<TargetFrameworkVersion>v4.0</TargetFrameworkVersion>";
            position = fileContent.IndexOf(net4);
            if (position > -1)
            {
                fileContent = fileContent.Replace(net4, "<TargetFrameworkVersion>v" + toNetVersion + "</TargetFrameworkVersion>");
                return;
            }

            string net5 = "<TargetFrameworkVersion>v4.5</TargetFrameworkVersion>";
            position = fileContent.IndexOf(net5);
            if (position > -1)
            {
                fileContent = fileContent.Replace(net4, "<TargetFrameworkVersion>v" + toNetVersion + "</TargetFrameworkVersion>");
                return;
            }

        }

        private void ChangeFormatVersionInSolutionFile(ref string fileContent, string toNetVersion)
        {
            if (toNetVersion == "4.5")
            {
                string replacedString = "";
                replacedString = "Format Version 12.00";

                string tools1 = "Format Version 10.00"; // vs2008
                int position = fileContent.IndexOf(tools1);
                if (position > -1)
                {
                    fileContent = fileContent.Replace(tools1, replacedString);
                    return;
                }

                string tools2 = "Format Version 11.00"; //vs2010
                position = fileContent.IndexOf(tools2);
                if (position > -1)
                {
                    fileContent = fileContent.Replace(tools2, replacedString);
                    return;
                }
            }
            else
            {
                string replacedString = "";
                replacedString = "Format Version 11.00";

                string tools1 = "Format Version 10.00"; // vs2008
                int position = fileContent.IndexOf(tools1);
                if (position > -1)
                {
                    fileContent = fileContent.Replace(tools1, replacedString);
                    return;
                }

                string tools2 = "Format Version 12.00"; // vs2012
                position = fileContent.IndexOf(tools2);
                if (position > -1)
                {
                    fileContent = fileContent.Replace(tools2, replacedString);
                    return;
                }
            }
        }

        #endregion

        #region Trigger

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
                    if (checkBoxChangeKeyFiles.Checked)
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
                if (fileContent.IndexOf(".csproj") > -1)
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
                    if (csMarkerSource.EndsWith("0"))
                        csMarkerSource = csMarkerSource.Substring(0, csMarkerSource.Length - 1);
                    if (csMarkerTarget.EndsWith("0"))
                        csMarkerTarget = csMarkerTarget.Substring(0, csMarkerTarget.Length - 1);

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
                    if (vbMarkerSource.EndsWith("0"))
                        vbMarkerSource = vbMarkerSource.Substring(0, vbMarkerSource.Length - 1);
                    if (vbMarkerTarget.EndsWith("0"))
                        vbMarkerTarget = vbMarkerTarget.Substring(0, vbMarkerTarget.Length - 1);

                    if (fileContent.IndexOf(vbMarkerSource) > -1)
                    {
                        fileContent = fileContent.Replace(vbMarkerSource, vbMarkerTarget);
                        File.Delete(file);
                        File.WriteAllText(file, fileContent, Encoding.UTF8);
                    }
                }
            }
        }

        private void buttonChooseFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fdg = new FolderBrowserDialog();
            if (DialogResult.OK == fdg.ShowDialog(this))
                textBoxFolder.Text = fdg.SelectedPath;
        }

        private void buttonLoadConfig_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dlg = new OpenFileDialog();
                dlg.InitialDirectory = Application.StartupPath;
                dlg.Filter = "Xml Files(*.xml)|*.xml";
                if (dlg.ShowDialog(this) == DialogResult.Cancel)
                    return;

                string targetFolder = string.Empty;
                bool changeNetMarker = false;
                string from = string.Empty;
                string to = string.Empty;
                bool changeKeyFiles = false;
                string keyFilesFolder = string.Empty;

                ConfigManager.LoadConfigurationFromConfigFile(dlg.FileName, ref targetFolder, ref changeNetMarker, ref from, ref to, ref changeKeyFiles, ref keyFilesFolder);

                textBoxFolder.Text = targetFolder;
                checkBoxChangeNetMarker.Checked = changeNetMarker;
                comboBoxFromNetVersion.Text = from;
                comboBoxToNetVersion.Text = to;
                checkBoxChangeKeyFiles.Checked = changeKeyFiles;
                textBoxKeyFilesRootFolder.Text = keyFilesFolder;
            }
            catch (Exception exception)
            {
                ExceptionDisplayer.ShowException(this, exception);
            }
        }

        private void buttonSaveConfig_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog dlg = new SaveFileDialog();
                dlg.InitialDirectory = Application.StartupPath;
                dlg.Filter = "Xml Files(*.xml)|*.xml";
                if (dlg.ShowDialog(this) == DialogResult.Cancel)
                    return;

                ConfigManager.SaveConfigurationToXMLFile(dlg.FileName, textBoxFolder.Text, checkBoxChangeNetMarker.Checked, comboBoxFromNetVersion.Text, comboBoxToNetVersion.Text, checkBoxChangeKeyFiles.Checked, textBoxKeyFilesRootFolder.Text);
            }
            catch (Exception exception)
            {
                ExceptionDisplayer.ShowException(this, exception);
            }
        }

        #endregion
    }
}
