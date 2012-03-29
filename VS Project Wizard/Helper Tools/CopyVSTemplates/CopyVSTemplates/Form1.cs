using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Xml.Linq;
using System.IO;
using System.Text;
using System.Windows.Forms;
using ICSharpCode.SharpZipLib;
using ICSharpCode.SharpZipLib.Zip;

using System.Xml;

namespace CopyVSTemplates
{
    public partial class Form1 : Form
    {
        protected static string _assemblyFolder = "NetOffice VS Wizard\\NetOffice Assemblies";
        protected static string _settingsFolder = "NetOffice VS Wizard";

        public Form1()
        {
            InitializeComponent();
            // textBoxRootFolder.Text = Application.StartupPath;
            labelNetOfficeSourceFolder.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), _settingsFolder);
            textBoxRootFolder.Text = @"C:\NetOffice\VS Project Wizard\Visual Studio Templates";
            textBoxVSSourceFolder.Text = @"C:\NetOffice\VS Project Wizard\NetOffice VS Wizard Files";
        }

        private void buttonPerformAction_Click(object sender, EventArgs e)
        {
            try
            {
                textBoxLog.Text = string.Empty;

                string vs2008Path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Visual Studio 2008\Templates\ProjectTemplates\NetOffice";
                string vs2010Path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Visual Studio 2010\Templates\ProjectTemplates\NetOffice";

                string[] directories = GetDirectories(textBoxRootFolder.Text.Trim());
                
                foreach (string directory in directories)
                {
                    string archiveFullName = PerformFolderCompression(directory);
                    string archiveName = Path.GetFileName(archiveFullName);

                    if (checkBoxVS2008.Checked)
                    {
                        if (!Directory.Exists(vs2008Path))
                            Directory.CreateDirectory(vs2008Path);

                        string destFileName = Path.Combine(vs2008Path, archiveName);
                        if (File.Exists(destFileName))
                            File.Delete(destFileName);

                        File.Copy(archiveFullName, destFileName);
                        textBoxLog.AppendText(string.Format("Copy {0}{1}", destFileName, Environment.NewLine));
                    }

                    if (checkBoxVS2010.Checked)
                    {
                        if (!Directory.Exists(vs2010Path))
                            Directory.CreateDirectory(vs2010Path);

                        string destFileName = Path.Combine(vs2010Path, archiveName);
                        if (File.Exists(destFileName))
                            File.Delete(destFileName);

                        File.Copy(archiveFullName, destFileName);
                        textBoxLog.AppendText(string.Format("Copy {0}{1}", destFileName, Environment.NewLine));
                    }

                    File.Delete(archiveFullName);
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Ein Fehler ist aufgetreten.\r\n\r\n" + exception.Message);
            }
        }

        private static string GetUpperPath(string path)
        {
            int position = path.LastIndexOf("\\");
            return path.Substring(0, position);
        }

        private static string GetFolderName(string path)
        {
            int position = path.LastIndexOf("\\");
            return path.Substring(position+1);
        }

        private static string PerformFolderCompression(string path)
        {
            string rootPath = GetUpperPath(path);
            string name = GetFolderName(path) + ".zip";
            string archiveFullName = Path.Combine(rootPath, name);
            string[] filenames = Directory.GetFiles(path);

            if (File.Exists(archiveFullName))
                File.Delete(archiveFullName);

            new FastZip().CreateZip(archiveFullName, path, true, null, "-.svn");
          
            return archiveFullName;
        }

        private static string[] GetDirectories(string path)
        { 
            List<string> result = new List<string>();
            string[] directories = Directory.GetDirectories(path);
            foreach (string directory in directories)
	        {
                if (IncludesTemplate(directory))
                    result.Add(directory);
	       }
            return result.ToArray();
        }

        private static bool IncludesTemplate(string path)
        {
            string[] directories = Directory.GetFiles(path, "*.vstemplate");
            return (directories.Length > 0);
        }

        private void buttonChooseFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.SelectedPath = textBoxRootFolder.Text;
            if (DialogResult.OK == dialog.ShowDialog(this))
                textBoxRootFolder.Text = dialog.SelectedPath;
        }

        private void buttonOpen2008Folder_Click(object sender, EventArgs e)
        {
            string vs2008Path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Visual Studio 2008\Templates\ProjectTemplates";
            if (Directory.Exists(vs2008Path))
                System.Diagnostics.Process.Start(vs2008Path);
            else
                MessageBox.Show("Folder doesnt exists.");
                
        }

        private void buttonOpen2010Folder_Click(object sender, EventArgs e)
        {
            string vs2010Path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Visual Studio 2010\Templates\ProjectTemplates";
            if (Directory.Exists(vs2010Path))
                System.Diagnostics.Process.Start(vs2010Path);
            else
                MessageBox.Show("Folder doesnt exists.");
        }

        private void buttonValidateSourceFolder_Click(object sender, EventArgs e)
        {
            string validateSourceFolder = textBoxVSSourceFolder.Text.Trim();

            string sourceFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), _settingsFolder);
            if (!Directory.Exists(sourceFolder))
                Directory.CreateDirectory(sourceFolder);

            string settingsFilePath = Path.Combine(sourceFolder, "Settings.xml");
            if(!File.Exists(settingsFilePath))
            {
                XDocument document = new XDocument(new XElement("Settings",new XElement("Language", new XAttribute("LCID","1031"))));
                document.Save(settingsFilePath);
            }

            string assemblyFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), _assemblyFolder);
            if (!Directory.Exists(assemblyFolder))
                Directory.CreateDirectory(assemblyFolder);

            string[] apiFiles = new string[] { "LateBindingApi.Core", "OfficeApi", "ExcelApi", "WordApi", "OutlookApi", "PowerPointApi", "AccessApi", "VBIDEApi", "DAOApi", "ADODBApi", "OWC10Api", "MSDATASRCApi", "MSComctlLibApi" };

            string[] netRuntimes = new string[] { "2.0", "3.0", "3.5", "4.0" };
            foreach (string runtime in netRuntimes)
            {
                string folder = Path.Combine(assemblyFolder, runtime);
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);

                foreach (string apiFile in apiFiles)
                {
                    string destinationFile = Path.Combine(assemblyFolder, runtime);
                    destinationFile = Path.Combine(destinationFile, apiFile + ".dll");

                    if (!File.Exists(destinationFile))
                    {
                        string sourceFile = Path.Combine(validateSourceFolder, runtime);
                        sourceFile = Path.Combine(sourceFile, apiFile + ".dll");

                        File.Copy(sourceFile, destinationFile);
                    }
                }
            }

            string docuFilesFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), _assemblyFolder);
            docuFilesFolder = Path.Combine(docuFilesFolder, "DocuFiles");
            if (!Directory.Exists(docuFilesFolder))
                Directory.CreateDirectory(docuFilesFolder);


            foreach (string apiFile in apiFiles)
            {
                string destinationFile = Path.Combine(assemblyFolder, "DocuFiles");
                destinationFile = Path.Combine(destinationFile, apiFile + ".xml");
                if (!File.Exists(destinationFile))
                {
                    string sourceFile = Path.Combine(validateSourceFolder, "DocuFiles");
                    sourceFile = Path.Combine(sourceFile, apiFile + ".xml");
                    File.Copy(sourceFile, destinationFile);
                }
            }

            MessageBox.Show("Done!", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void buttonDeleteSourceFolder_Click(object sender, EventArgs e)
        {
            string validateSourceFolder = labelNetOfficeSourceFolder.Text.Trim();
            if (Directory.Exists(validateSourceFolder))
            {
                Directory.Delete(validateSourceFolder, true);
                MessageBox.Show("Done!", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("The folder doesnt exists.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

    }
}
