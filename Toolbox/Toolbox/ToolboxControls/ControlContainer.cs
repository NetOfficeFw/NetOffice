using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NetOffice.DeveloperToolbox.Controls.InfoLayer;

namespace NetOffice.DeveloperToolbox.ToolboxControls
{
    /// <summary>
    /// Wraps a toolbox control instance as a proxy to communicate between host and toolbox control instance
    /// </summary>
    public partial class ControlContainer : UserControl, IToolboxControl
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public ControlContainer()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="innerInstance">instance to wrap into</param>
        public ControlContainer(IToolboxControl innerInstance)
        {
            InitializeComponent();
            if (null == innerInstance)
                throw new ArgumentNullException("innerInstance");
            InnerInstance = innerInstance;
            panelToolboxControl.Controls.Add(innerInstance as Control);
            (innerInstance as Control).Dock = DockStyle.Fill;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Inner(real) toolbox instance 
        /// </summary>
        internal IToolboxControl InnerInstance { get; private set; }

        #endregion

        #region IToolboxControl

        public IToolboxHost Host
        {
            get
            {
                return InnerInstance.Host;
            }
        }

        public string ControlName
        {
            get { return InnerInstance.ControlName; }
        }

        public string ControlCaption
        {
            get { return InnerInstance.ControlCaption; }
        }

        public Image Icon
        {
            get { return InnerInstance.Icon; }
        }

        public bool SupportsHelpContent
        {
            get 
            {
                return InnerInstance.SupportsHelpContent;
            }
        }

        public bool SupportsInfoMessage
        {
            get
            {
                return InnerInstance.SupportsInfoMessage;
            }
        }

        public ToolboxControlMessageKind InfoMessageKind
        {
            get
            {
                return InnerInstance.InfoMessageKind;
            }
        }

        public string InfoMessage
        {
            get
            {
                return InnerInstance.InfoMessage;
            }
        }

        public void InitializeControl(IToolboxHost host)
        {
            InnerInstance.InitializeControl(host);
        }

        public void Activate(bool firstTime)
        {
            buttonInfo.Visible = InnerInstance.SupportsHelpContent;
            InnerInstance.Activate(firstTime);
            if (InnerInstance.SupportsHelpContent)
                controlBackColorAnimator1.Start(false);
            SetupInfoMessage();
        }

        public void Deactivated()
        {
            InnerInstance.Deactivated();
            if (InnerInstance.SupportsHelpContent)
                controlBackColorAnimator1.Start(false);
        }

        public void LoadComplete()
        {
            InnerInstance.LoadComplete();
        }

        public void LoadConfiguration(System.Xml.XmlNode configNode)
        {
            InnerInstance.LoadConfiguration(configNode);
        }

        public void SaveConfiguration(System.Xml.XmlNode configNode)
        {
            InnerInstance.SaveConfiguration(configNode);
        }

        public void SetLanguage(int id)
        {
            Translation.ToolLanguage language =  Host.Languages.Where(l => l.LCID == id).FirstOrDefault();
            if (null != language)
            {
                string space = InnerInstance.ControlName.Substring(0, InnerInstance.ControlName.IndexOf("."));
                var component = language.Components[space];
                Translation.Translator.TranslateControls(InnerInstance as Control, component.ControlRessources);
            }
            else 
            {
                string space = InnerInstance.ControlName.Substring(0, InnerInstance.ControlName.IndexOf("."));
                Translation.Translator.TranslateControls(InnerInstance as Control, String.Format("ToolboxControls.{0}.Strings.txt", space), id);
            }

            SetupInfoMessage();
        }

        public Stream GetHelpText(int lcid)
        {
            return InnerInstance.GetHelpText(lcid);
        }

        public new void KeyDown(KeyEventArgs e)
        {
            InnerInstance.KeyDown(e);
        }

        public void Release()
        {
            InnerInstance.Release();
        }

        public IContainer Components
        {
            get { return InnerInstance.Components; }
        }

        #endregion

        #region ILocalizationDesign

        public void EnableDesignView(int lcid, string parentComponentName)
        {
            InnerInstance.EnableDesignView(lcid, parentComponentName);
        }

        public void Localize(Translation.ItemCollection strings)
        {
            InnerInstance.Localize(strings);
        }

        public void Localize(string name, string text)
        {
            InnerInstance.Localize(name, text);
        }

        public string GetCurrentText(string name)
        {
            return InnerInstance.GetCurrentText(name);
        }

        public string NameLocalization
        {
            get
            {
                return InnerInstance.NameLocalization;
            }
        }

        public IEnumerable<ILocalizationChildInfo> Childs
        {
            get
            {
                return InnerInstance.Childs;
            }
        }

        #endregion

        #region Methods

        private void SetupInfoMessage()
        {
            if (InnerInstance.SupportsInfoMessage)
            {
                switch (InnerInstance.InfoMessageKind)
                {
                    case ToolboxControlMessageKind.Information:
                        pictureBoxInformation.Visible = true;
                        pictureBoxWarning.Visible = false;
                        break;
                    case ToolboxControlMessageKind.Warning:
                        pictureBoxInformation.Visible = false;
                        pictureBoxWarning.Visible = true;
                        break;
                    default:
                        pictureBoxInformation.Visible = false;
                        pictureBoxWarning.Visible = false;
                        break;
                }
                labelInfoMessage.Text = InnerInstance.InfoMessage;
                labelInfoMessage.Visible = true;
            }
            else
            {
                labelInfoMessage.Text = String.Empty;
                pictureBoxInformation.Visible = false;
                pictureBoxWarning.Visible = false;
                labelInfoMessage.Visible = false;
            }
        }

        #endregion

        #region Trigger

        private void buttonInfo_Click(object sender, EventArgs e)
        {
            try
            {
                Stream stream = InnerInstance.GetHelpText(InnerInstance.Host.CurrentLanguageID);
                InfoControl infoBox = new InfoControl(stream);
                this.Controls.Add(infoBox);
                infoBox.BringToFront();
                infoBox.Show();
                stream.Close();
                stream.Dispose();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception,ErrorCategory.NonCritical, InnerInstance.Host.CurrentLanguageID);
            }
        }

        #endregion
    }
}
