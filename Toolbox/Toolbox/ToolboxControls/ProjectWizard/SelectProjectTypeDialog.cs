using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Security.Principal;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Win32;

namespace NetOffice.DeveloperToolbox
{
    partial class SelectProjectTypeDialog : Form
    {
        #region .ctor()

        public SelectProjectTypeDialog()
        {
            InitializeComponent();
            comboBoxNetRuntime.SelectedIndex = 2;
            if (IsAdministrator())
            {
                labelNoAdminHint.Visible = false;
                radioButtonDesktop.Enabled = true;
                radioButtonUserFolder.Enabled = true;
                radioButtonVSProjectFolder.Enabled = true;
            }
            Translator.TranslateControls(this, "ProjectWizard.SelectProjectTypeDialog.txt", ProjectWizardControl.CurrentLanguageID);
        }

        #endregion

        #region Properties

        public ProjectOptions SelectedOptions
        {
            get
            {
                ProjectOptions options = new ProjectOptions(GetSelectedFolder(),
                                                           1.0,
                                                            GetSelectedProjectType(),
                                                            GetSelectedLanguage(),
                                                            GetSelectedIDE()
                                                            );
                return options;
            }
        }

        #endregion

        #region Trigger

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        private void comboBoxNetRuntime_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxNetRuntime.SelectedIndex == 3)
            {
                radioButtonVS2008.Enabled = false;
                radioButtonVS2010.Checked = true;
            }
            else
            {
                radioButtonVS2008.Enabled = true;
            }
        }

        private void buttonSelect_Click(object sender, EventArgs e)
        {
            if (radioButtonCustomFolder.Checked)
            {
                if (textBoxCustomFolder.Text == "")
                {
                    string message = ProjectWizardControl.CurrentLanguageID == 1031 ? "Bitte wählen Sie einen benutzerdefinierten Speicherordner." : "Choose a custom folder first.";
                    MessageBox.Show(this, message, "Developer Toolbox", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    return;
                }

                if (!IsAdministrator())
                {
                    string message1 =  ProjectWizardControl.CurrentLanguageID == 1031 ? "Sie haben einen benutzerdefinierten Ordner ausgewählt und die Developer Toolbox verfügt derzeit nicht über erhöhte Berechtigungen." : "You have a custom folder selected and the Toolbox missed Administratr Privileges.";
                    string message2 =  ProjectWizardControl.CurrentLanguageID == 1031 ?"Soll die Developer Toolbox den Schreibzugriff jetzt prüfen um später mögliche Fehler zu vermeiden?" : "The Developer Toolbox want to do a test access to avoid any problems. Its okay?";
                    if (DialogResult.Yes == MessageBox.Show(this, string.Format("{0}{2}{2}{1}", message1, message2, Environment.NewLine), "Developer Toolbox", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
                    {
                        if (DoCustomTestWrite())
                            MessageBox.Show(this,  ProjectWizardControl.CurrentLanguageID == 1031 ? "Die Schreibprüfung verlief erfolgreich." : "Already was okay", "Developer Toolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        else
                        {
                            string message = ProjectWizardControl.CurrentLanguageID == 1031 ? "Die Schreibprüfung ist geschlagen." + Environment.NewLine + Environment.NewLine + "Bitte wählen Sie einen anderen Speicherordner" :
                                "The test failed." + Environment.NewLine + Environment.NewLine + "Please choose a different folder or run Developer Toolbox with Administrator Privileges.";
                            MessageBox.Show(this, message, "Developer Toolbox", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
            }

            if (radioButtonVSProjectFolder.Checked)
            {
                if (GetVisualStudioExpressProjectFolder() == null)
                {
                    string target = (radioButtonCSharp.Checked ? "VCS" : "VB") + "Express " + (radioButtonVS2008.Checked ? "2008" : "2010");
                    string messageDE = string.Format("Developer Toolbox konnte den Projektordner für {0} nicht ermitteln.\r\n\r\nBitte wählen Sie einen anderen Ordner.", target);
                    string messageEN = string.Format("Developer Toolbox failed to detect the {0} project folder.\r\n\r\nPlease choose a different folder.", target);
                    MessageBox.Show(this, ProjectWizardControl.CurrentLanguageID == 1031 ? messageDE : messageEN, "Developer Toolbox", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }
        
        private void radioButtonCustomFolder_CheckedChanged(object sender, EventArgs e)
        {
            buttonChooseFolder.Enabled = radioButtonCustomFolder.Checked;
            if (textBoxCustomFolder.Text == "" && radioButtonCustomFolder.Checked)
                buttonChooseFolder_Click(buttonChooseFolder, new EventArgs());
        }

        private void buttonChooseFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (DialogResult.OK == dialog.ShowDialog(this))
                textBoxCustomFolder.Text = dialog.SelectedPath;
        }

        #endregion

        #region Methods

        private IDE GetSelectedIDE()
        {
            if (radioButtonVS2008.Checked)
                return IDE.VS2008;
            else
                return IDE.VS2010;
        }

        private ProgrammingLanguage GetSelectedLanguage()
        {
            if (radioButtonCSharp.Checked)
                return ProgrammingLanguage.CSharp;
            else
                return ProgrammingLanguage.VB;
        }

        private ProjectType GetSelectedProjectType()
        {
            if (radioButtonAutomationAddin.Checked)
                return ProjectType.Addin;

            if (radioButtonWindowsForms.Checked)
                return ProjectType.WindowsForms;

            if (radioButtonClassLibrary.Checked)
                return ProjectType.ClassLibrary;

            return ProjectType.Console;
        }

        private string GetSelectedFolder()
        {
            if (radioButtonCustomFolder.Checked)
                return textBoxCustomFolder.Text;

            if (radioButtonDesktop.Checked)
                return Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            if (radioButtonUserFolder.Checked)
                return Environment.GetFolderPath(Environment.SpecialFolder.Personal);

            if (radioButtonVSProjectFolder.Checked)
            {
                string visualStudioProjects = GetVisualStudioExpressProjectFolder();
                return visualStudioProjects;
            }

            return Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

        }

        private string GetVisualStudioInstallFolder()
        {
            string result = null;
            if (radioButtonCSharp.Checked)
            {
                string path = "-1";
                string folderExpress = "Software\\Microsoft\\VCSExpress\\" + (radioButtonVS2008.Checked == true ? "9.0_Config" : "10.0_Config");
                RegistryKey key = Registry.CurrentUser.OpenSubKey(folderExpress, false);
                if (null != key)
                {
                    path = key.GetValue("InstallDir", "-1") as string;
                    key.Close();
                }

                if ("-1" != path)
                {
                    result = path;
                }
            }
            else
            {
                string path = "-1";
                string folderExpress = "Software\\Microsoft\\VBExpress\\" + (radioButtonVS2008.Checked == true ? "9.0_Config" : "10.0_Config");
                RegistryKey key = Registry.CurrentUser.OpenSubKey(folderExpress, false);
                if (null != key)
                {
                    path = key.GetValue("InstallDir", "-1") as string;
                    key.Close();
                }

                if ("-1" != path)
                {
                    result = path;
                }
            }

            return result;
        }

        private string GetVisualStudioExpressProjectFolder()
        {
            string result = null;
            if (radioButtonCSharp.Checked)
            {
                string path = "-1";
                string folderExpress = "Software\\Microsoft\\VCSExpress\\" + (radioButtonVS2008.Checked == true ? "9.0" : "10.0");
                RegistryKey key = Registry.CurrentUser.OpenSubKey(folderExpress, false);
                if (null != key)
                {
                    path = key.GetValue("VisualStudioProjectsLocation", "-1") as string;
                    key.Close();
                }

                if ("-1" != path)
                {
                    path = path.Replace("%USERPROFILE%", Environment.GetFolderPath(Environment.SpecialFolder.UserProfile));
                    path = path.Replace("%APPDATA%", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData));
                    path = path.Replace("%PROGRAMFILES%", Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles));
                    path = path.Replace("%COMMONPROGRAMFILES%", Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles));
                    path = path.Replace("%VsInstallDir%", GetVisualStudioInstallFolder());
                    result = path;
                }
            }
            else
            {
                string path = "-1";
                string folderExpress = "Software\\Microsoft\\VBExpress\\" + (radioButtonVS2008.Checked == true ? "9.0" : "10.0");
                RegistryKey key = Registry.CurrentUser.OpenSubKey(folderExpress, false);
                if (null != key)
                {
                    path = key.GetValue("VisualStudioProjectsLocation", "-1") as string;
                    key.Close();
                }

                if ("-1" != path)
                {
                    path = path.Replace("%USERPROFILE%", Environment.GetFolderPath(Environment.SpecialFolder.UserProfile));
                    path = path.Replace("%APPDATA%", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData));
                    path = path.Replace("%PROGRAMFILES%", Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles));
                    path = path.Replace("%COMMONPROGRAMFILES%", Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles));
                    path = path.Replace("%VsInstallDir%", GetVisualStudioInstallFolder());
                    result = path;
                }
            }

            return result;
        }

        private bool DoCustomTestWrite()
        {
            string folderPath = textBoxCustomFolder.Text;
            if (!System.IO.Directory.Exists(folderPath))
                CreateDirectory(folderPath);

            return CreateTestFile(folderPath);
        }

        private bool CreateTestFile(string folderPath)
        {
            string testFile = System.IO.Path.Combine(folderPath, Guid.NewGuid().ToString() + ".txt");
            try
            {
                System.IO.File.AppendAllText(testFile, DateTime.Now.ToString());
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (System.IO.File.Exists(testFile))
                    System.IO.File.Delete(testFile);
            }
        }

        private bool CreateDirectory(string folderPath)
        {
            try
            {
                System.IO.Directory.CreateDirectory(folderPath);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool IsAdministrator()
        {
            WindowsIdentity myWindowsIdentity = WindowsIdentity.GetCurrent();
            WindowsPrincipal myWindowsPrincipal = new WindowsPrincipal(myWindowsIdentity);
            return myWindowsPrincipal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        #endregion

        private void radioButtonApplicationData_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
