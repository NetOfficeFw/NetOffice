using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Mono.Cecil;

namespace NetOffice.DeveloperToolbox.ToolboxControls.OfficeCompatibility
{
    /// <summary>
    /// Analyze assemblies for NetOffice requests to check how compatible is the solution
    /// </summary>
    [RessourceTable("ToolboxControls.OfficeCompatibility.Strings.txt")]
    public partial class OfficeCompatibilityControl : UserControl, IToolboxControl
    {
        #region Fields
     
        private AnalyzerResult _result;
        private string _assemblyFullFileName;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public OfficeCompatibilityControl()
        {
            InitializeComponent();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Last Right Click Link Label
        /// </summary>
        private LinkLabel LastClickedLinkLabel { get; set; }

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
            PictureBox box16 = tableLayoutResult.Controls["pictureBox" + name + "16"] as PictureBox;

            if (name != "Project" && name != "Visio")
            {
                SetImage(box09, info[0].Support);
                SetImage(box10, info[1].Support);
            }
            SetImage(box11, info[2].Support);
            SetImage(box12, info[3].Support);
            SetImage(box14, info[4].Support);
            SetImage(box15, info[5].Support);
            SetImage(box16, info[6].Support);
        }

        private void ShowBadImageFormatError()
        {
            labelErrorMessage.Text = labelBadImageError.Text;
            panelAssemblyError.Visible = true;
        }

        private void ShowNoNetOfficeError()
        {
            labelErrorMessage.Text = labelNoNetOfficeError.Text;
            panelAssemblyError.Visible = true;
        }

        private void ShowUnexpectedError(Exception exception)
        {
            labelErrorMessage.Text = exception.Message;
            panelAssemblyError.Visible = true;
        }

        private void HideError()
        {
            panelAssemblyError.Visible = false;
        }

        #endregion

        #region IToolboxControl

        public IToolboxHost Host { get; private set; }

        public new void KeyDown(KeyEventArgs e)
        { 
            
        }

        public string ControlName
        {
            get { return "OfficeCompatibility.OfficeCompatibilityControl"; }
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
            get { return Ressources.RessourceUtils.ReadImageFromRessource("ToolboxControls.OfficeCompatibility.Icon.png"); }
        }

        public bool SupportsHelpContent
        {
            get
            {
                return true;
            }
        }

        public bool SupportsInfoMessage
        {
            get
            {
                return true;
            }
        }

        public ToolboxControlMessageKind InfoMessageKind
        {
            get
            {
                return ToolboxControlMessageKind.Information;
            }
        }

        public string InfoMessage
        {
            get
            {
                return labelDebugHint.Text;
            }
        }

        public void LoadComplete()
        {
            
        }

        public void Activate(bool firstTime)
        {

        }

        public void Deactivated()
        {

        }

        public void LoadConfiguration(System.Xml.XmlNode configNode)
        {
            
        }

        public void SaveConfiguration(System.Xml.XmlNode configNode)
        {
            
        }

        public void InitializeControl(IToolboxHost host)
        {
            Host = host;
        }

        public Stream GetHelpText()
        {
            return Ressources.RessourceUtils.ReadStream("ToolboxControls.OfficeCompatibility.Info1033.rtf");
        }

        public void Release()
        {
           
        }

        #endregion
  
        #region Trigger
        
        private void buttonSelectAssembly_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "*.dll|*.dll|*.exe|*.exe|All Files|*.*";        
                if (DialogResult.OK  != dialog.ShowDialog(this))
                    return;

                _assemblyFullFileName = dialog.FileName;
                buttonRefresh_Click(this, new EventArgs());
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {
            try
            {
                if (null == _result)
                    return;

                ReportControl reportBox = new ReportControl(_result);
                this.Controls.Add(reportBox);
                reportBox.Dock = DockStyle.Fill;
                reportBox.BringToFront();
                reportBox.Show();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical);
            }
        }

        private void buttonRefresh_Click(object sender, EventArgs e)
        {
            try
            {
                HideError();

                AssemblyDefinition assemblyDefinition = AssemblyDefinition.ReadAssembly(_assemblyFullFileName);
                textBoxAssembly.Text = assemblyDefinition.Name.ToString();

                _result = AssemblyAnalyzer.AnalyzeAssembly(assemblyDefinition);
                if (_result.ContainsNetOfficeReferences)
                {
                    buttonRefresh.Enabled = true;
                    buttonReport.Enabled = true;
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
                    ShowNoNetOfficeError();
                }
            }
            catch (BadImageFormatException)
            {
                    buttonRefresh.Enabled = false;
                    buttonReport.Enabled = false;
                    ShowBadImageFormatError();
            }
            catch (Exception exception)
            {
                ShowUnexpectedError(exception);
            }
        }

        private void linkLabelNotSupported_Click(object sender, EventArgs args)
        {
            try
            {
                MouseEventArgs mouseArgs = args as MouseEventArgs;
                if (null == mouseArgs)
                    return;

                if (mouseArgs.Button == MouseButtons.Left)
                {
                    LinkLabel label = sender as LinkLabel;
                    System.Diagnostics.Process.Start(label.Tag as string);
                }
                else if (mouseArgs.Button == MouseButtons.Right)
                {
                    LastClickedLinkLabel = sender as LinkLabel;
                    LinkContextMenu.Show(sender as Control, 0, 0);
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void LinkContextMenu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (null != LastClickedLinkLabel)
            {
                Clipboard.SetText(LastClickedLinkLabel.Tag as string);
            }
        }

        #endregion
    }
}