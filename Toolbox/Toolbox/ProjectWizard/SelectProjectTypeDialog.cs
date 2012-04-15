using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Security.Principal;
using System.Reflection;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox
{
    partial class SelectProjectTypeDialog : Form
    {
        #region .ctor()

        public SelectProjectTypeDialog()
        {
            InitializeComponent();
            comboBoxNetRuntime.SelectedIndex = 3;
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
                                                            Convert.ToDouble(comboBoxNetRuntime.Text),
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
                    MessageBox.Show(this, "Bitte wählen Sie einen benutzerdefinierten Speicherordner.", "Developer Toolbox", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    return;
                }

                if (!IsAdministrator())
                {
                    string message1 = "Sie haben einen benutzerdefinierten Ordner ausgewählt und die Developer Toolbox verfügt derzeit nicht über erhöhte Berechtigungen.";
                    string message2 = "Soll die Developer Toolbox den Schreibzugriff jetzt prüfen um später mögliche Fehler zu vermeiden?";
                    if (DialogResult.Yes == MessageBox.Show(this, string.Format("{0}{2}{2}{1}", message1, message2, Environment.NewLine), "Developer Toolbox", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
                    {
                        if (DoCustomTestWrite())
                            MessageBox.Show(this, "Die Schreibprüfung verlief erfolgreich.", "Developer Toolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        else
                        {
                            MessageBox.Show(this, "Die Schreibprüfung ist geschlagen." + Environment.NewLine + Environment.NewLine +"Bitte wählen Sie einen anderen Speicherordner", "Developer Toolbox", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
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
                string visualStudio = "";
                if(radioButtonVS2008.Checked)
                    visualStudio = "VisualStudio\\Projects";
                else 
                    visualStudio = "VisualStudio\\Projects";

                return System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), visualStudio);
            }

            return Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

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
    }
}
