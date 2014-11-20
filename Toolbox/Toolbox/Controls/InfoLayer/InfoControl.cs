using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using  System.Reflection;
using System.IO;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Controls.InfoLayer
{
    [RessourceTable("Controls.InfoLayer.Strings.txt")]
    public partial class InfoControl : UserControl, ILocalizationDesign, ILocalizationReplaceProvider
    {
        #region Fields

        private int _designLCID;
        private string _parentComponentName;

        #endregion

        #region Ctor

        public InfoControl()
        {
            InitializeComponent();
        }

        public InfoControl(string text)
        {
            InitializeComponent();
            this.Dock = DockStyle.Fill;
            richTextBoxHelpContent.Text = text;            
        }

        public InfoControl(Stream rtfStream)
        {
            InitializeComponent();
            this.Dock = DockStyle.Fill;
            richTextBoxHelpContent.LoadFile(rtfStream, RichTextBoxStreamType.RichText);
        }

        public InfoControl(string text, bool isRessourceAddress)
        {
            InitializeComponent();
            this.Dock = DockStyle.Fill;

            if (isRessourceAddress)
            {
                richTextBoxHelpContent.LoadFile(ReadStream(text), RichTextBoxStreamType.RichText);
            }
            else
            {
                richTextBoxHelpContent.Text = text;
            }
        }

        #endregion

        #region Trigger

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private static Stream ReadStream(string resId)
        {
            Assembly ass = Assembly.GetExecutingAssembly();
            string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            System.IO.Stream ressourceStream = ass.GetManifestResourceStream(assemblyName + "." + resId);
            if (ressourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));
            return ressourceStream;
        }

        private void richTextBox_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(e.LinkText);
            }
            catch
            {
                ;
            }
        }

        #endregion

        #region ILocalizationDesign

        public void EnableDesignView(int lcid, string parentComponentName)
        {
            _designLCID = lcid;
            _parentComponentName = parentComponentName;
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

        public IContainer Components
        {
            get { return components; }
        }

        public string NameLocalization
        {
            get { throw new NotImplementedException(); }
        }

        public IEnumerable<ILocalizationChildInfo> Childs
        {
            get { throw new NotImplementedException(); }
        }

        #endregion

        #region ILocalizationReplaceProvider
        
        public string Replace(string marker)
        {
            if (marker == "{0:$HelpContent}")
            {
                string target = _parentComponentName.Substring(0, _parentComponentName.LastIndexOf(".")) + ".Info" + _designLCID  + ".rtf";
                string content =Ressources.RessourceUtils.ReadString(target, false, false);
                return content;
            }
            else
                return "";
        }

        #endregion
    }
}
