using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Linq;
using System.Xml.Linq;

using Mono.Cecil;
using Mono.Cecil.Cil;

namespace NetOffice.DeveloperToolbox.OfficeCompatibility
{
    public partial class OfficeCompatibilityControl : UserControl, IToolboxControl
    {
        #region Fields

        int _currentLanguageID = 1031;
        AnalyzerResult _result;
        string _assemblyFullFileName;

        #endregion

        #region Construction

        public OfficeCompatibilityControl()
        {
            InitializeComponent();
        }
               
        #endregion

        #region IUtilsControl Members

        public string ControlName
        {
            get { return "OfficeCompatibility"; }
        }

        public string ControlCaption
        {
            get { return "Office Compatibility"; }
        }

        public System.ComponentModel.IContainer Components
        {
            get
            {
                return components;
            }
        }

        public Image Icon
        {
            get
            {
                return ReadImageFromRessource("Icon.png");
            }
        }

        public void LoadComplete()
        {

        }

        public void Activate()
        {
           
        }

        public void LoadConfiguration(System.Xml.XmlNode configNode)
        {
            
        }

        public void SaveConfiguration(System.Xml.XmlNode configNode)
        {
            
        }

        public void SetLanguage(int id)
        {
            _currentLanguageID = id;
            Translator.TranslateControls(this, "OfficeCompatibility.MessageTable.txt", id);
        }

        public void Release()
        {
           
        }

        #endregion

        #region Gui Trigger
        
        private void buttonSelectAssembly_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "*.exe|*.exe|*.dll|*.dll|All Files|*.*";        
                if (DialogResult.OK  != dialog.ShowDialog(this))
                    return;

                _assemblyFullFileName = dialog.FileName;
                buttonRefresh_Click(this, new EventArgs());
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void buttonInfo_Click(object sender, EventArgs e)
        {
            try
            {
                InfoControl infoBox = new InfoControl("OfficeCompatibility.Info" + _currentLanguageID.ToString() + ".rtf", true);
                this.Controls.Add(infoBox);
                infoBox.BringToFront();
                infoBox.Show();
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {
            try
            {
                if (null == _result)
                    return;

                ReportControl reportBox = new ReportControl(_result, _currentLanguageID);
                this.Controls.Add(reportBox);
                reportBox.Dock = DockStyle.Fill;
                reportBox.BringToFront();
                reportBox.Show();
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void buttonRefresh_Click(object sender, EventArgs e)
        {
            try
            {
                AssemblyDefinition assemblyDefinition = AssemblyDefinition.ReadAssembly(_assemblyFullFileName);
                textBoxAssembly.Text = assemblyDefinition.Name.ToString();

                _result = AssemblyAnalyzer.AnalyzeAssembly(assemblyDefinition);
                if (_result.ContainsNetOfficeReferences)
                {
                    buttonRefresh.Enabled = true;
                    buttonReport.Enabled = true;
                    panelInvalidAssembly.Visible = false;
                    panelVersionTable.Visible = true;
                    panelNoNetOfficeReferences.Visible = false;
                    panelResultTable.Visible = true;
                    SetupVersionInfo(_result.Office, "Office");
                    SetupVersionInfo(_result.Excel, "Excel");
                    SetupVersionInfo(_result.Word, "Word");
                    SetupVersionInfo(_result.Outlook, "Outlook");
                    SetupVersionInfo(_result.PowerPoint, "PowerPoint");
                    SetupVersionInfo(_result.Access, "Access");
                    SetupVersionInfo(_result.Project, "Project");
                    SetupVersionInfo(_result.Visio, "Visio");
                }
                else
                {
                    buttonRefresh.Enabled = false;
                    buttonReport.Enabled = false;
                    panelInvalidAssembly.Visible = false;
                    panelVersionTable.Visible = false;
                    panelNoNetOfficeReferences.Visible = true;
                    panelResultTable.Visible = false;
                }
            }
            catch (BadImageFormatException)
            {
                    buttonRefresh.Enabled = false;
                    buttonReport.Enabled = false;
                    panelVersionTable.Visible = false;
                    panelNoNetOfficeReferences.Visible = false;
                    panelResultTable.Visible = false;
                    panelInvalidAssembly.Visible = true;
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }

        private void linkLabelNotSupported_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                if (_currentLanguageID == 1031)
                    System.Diagnostics.Process.Start("http://netoffice.codeplex.com/wikipage?title=UnsupportedVersions_DE");
                else
                    System.Diagnostics.Process.Start("http://netoffice.codeplex.com/wikipage?title=UnsupportedVersions_EN");
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }
        
        #endregion

        #region Methods

        private void SetImage(PictureBox box, SupportVersion version)
        {
            switch (version)
            {
                case SupportVersion.Support:
                    box.BackgroundImage = pictureBoxOk.Image;
                    break;
                case SupportVersion.NotSupport:
                    box.BackgroundImage = pictureBoxProblem.Image;
                    break;
                default:
                    box.BackgroundImage = null;
                    break;
            }
        }

        private void SetupVersionInfo(SupportInfo[] info, string name)
        {
            PictureBox box09 = tableLayoutResult.Controls["pictureBox" + name + "09"] as PictureBox;
            PictureBox box10 = tableLayoutResult.Controls["pictureBox" + name + "10"] as PictureBox;
            PictureBox box11 = tableLayoutResult.Controls["pictureBox" + name + "11"] as PictureBox;
            PictureBox box12 = tableLayoutResult.Controls["pictureBox" + name + "12"] as PictureBox;
            PictureBox box14 = tableLayoutResult.Controls["pictureBox" + name + "14"] as PictureBox;
            PictureBox box15 = tableLayoutResult.Controls["pictureBox" + name + "15"] as PictureBox;

            if (name != "Project" && name != "Visio")
            { 
                SetImage(box09, info[0].Support);
                SetImage(box10, info[1].Support);
            }
            SetImage(box11, info[2].Support);
            SetImage(box12, info[3].Support);
            SetImage(box14, info[4].Support);
            SetImage(box15, info[5].Support);
        }

        #endregion

        #region Static Methods

        private static Image ReadImageFromRessource(string ressourcePath)
        {
            System.IO.Stream ressourceStream = null;
            string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            ressourcePath = assemblyName + ".OfficeCompatibility." + ressourcePath;
            ressourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(ressourcePath);
            if (ressourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));
            Bitmap newIcon = new Bitmap(ressourceStream);
            return newIcon;
        }

        #endregion
    }
}
