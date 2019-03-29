using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.About
{
    /// <summary>
    /// Application about panel
    /// </summary>
    [RessourceTable("ToolboxControls.About.Strings.txt")]
    public partial class AboutControl : UserControl, IToolboxControl
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public AboutControl()
        {
            InitializeComponent();
            labelVersionText.Text = String.Format("Version {0}", AssemblyInfo.AssemblyVersion);
            labelCopyrightText.Text = AssemblyInfo.AssemblyCopyright;
        }

        #endregion

        #region IToolboxControl

        public IToolboxHost Host { get; private set; }

        public string ControlName
        {
            get { return "About.AboutControl"; }
        }

        public string ControlCaption
        {
            get { return "About"; }
        }

        public Image Icon
        {
            get { return Ressources.RessourceUtils.ReadImageFromRessource("ToolboxControls.About.info_rhombus.png"); }
        }

        public bool SupportsHelpContent 
        {
            get
            {
                return false;
            }
        }

        public bool SupportsInfoMessage
        {
            get
            {
                return false;
            }
        }

        public ToolboxControlMessageKind InfoMessageKind
        {
            get 
            {
                return ToolboxControlMessageKind.Uncategorized;
            }
        }

        public string InfoMessage
        {
            get 
            {
                return String.Empty;
            }
        }

        public void InitializeControl(IToolboxHost host)
        {
            Host = host;
        }

        public void Activate(bool firstTime)
        {
            scroller1.TextToScroll = GetLanguageCredits();
            scroller1.Start();
            controlForeColorAnimator1.Start(false);
        }

        public void Deactivated()
        {
            scroller1.Stop();
            controlForeColorAnimator1.Stop();
        }

        public void LoadComplete()
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
            
        }

        public Stream GetHelpText(int lcid)
        { 
            throw new NotImplementedException();
        }

        public new void KeyDown(KeyEventArgs e)
        {
            
        }

        public void Release()
        {
            
        }

        public IContainer Components
        {
            get { return components; }
        }

        #endregion

        #region ILocalizationDesign

        public void EnableDesignView(int lcid, string parentComponentName)
        {

        }

        public void Localize(Translation.ItemCollection strings)
        {
            Translation.Translator.TranslateControls(this, strings);
        }

        public void Localize(string name, string text)
        {
            Translation.Translator.TranslateControl(this, name, text);
        }

        public string GetCurrentText(string name)
        {
            return Translation.Translator.TryGetControlText(this, name);
        }

        public string NameLocalization
        {
            get
            {
                return null;
            }
        }

        public IEnumerable<ILocalizationChildInfo> Childs
        {
            get
            {
                return new ILocalizationChildInfo[0];
            }
        }

        #endregion

        #region Methods

        private string GetLanguageCredits()
        {
            StringBuilder sb = new StringBuilder();
            foreach (var item in Host.Languages)
            {
                if (!String.IsNullOrWhiteSpace(item.Author))
                {
                    string lng = String.Format("{0}{1}{2}{3}{4}{4}{4}",
                        item.DisplayName + Environment.NewLine + Environment.NewLine, 
                        item.Author + Environment.NewLine, 
                        String.IsNullOrWhiteSpace(item.AuthorMail) ? "" : "   " + item.AuthorMail + Environment.NewLine,
                        String.IsNullOrWhiteSpace(item.AuthorSite) ? "" : "   " + item.AuthorSite + Environment.NewLine,
                        Environment.NewLine);
                    sb.Append(lng);
                }
            }
            return sb.ToString();
        }

        #endregion

        #region Trigger

        private void AboutControl_Resize(object sender, EventArgs e)
        {
            try
            {
                panelMain.Location = new Point((this.Width / 2) - (panelMain.Width / 2), (this.Height / 2) - (panelMain.Height / 2));
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }

        private void LinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                LinkLabel label = sender as LinkLabel;
                System.Diagnostics.Process.Start(label.Text);
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical, Host.CurrentLanguageID);
            }
        }
        #endregion
    }
}
