using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Xml;
using System.Reflection;
using System.Windows.Forms;

namespace TutorialsBase
{
    partial class FormOptions : TutorialForm
    {
        static string _fullConfigFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) , "NetOfficeTutorialsCS4.xml");
        static int _lcid = DefaultLCID;
        static bool _connectToDocumentation = true;
        static bool _saveSettings = true;

        public FormOptions()
        {
            InitializeComponent();

            if (_lcid != 1033)
                radioButtonLanguage1031.Checked = true;

            if (!_connectToDocumentation)
                radioButtonShowLink.Checked = true;

            if (!_saveSettings)
                checkBoxSaveSettings.Checked = true;
        }

        #region Methods

        public static void LoadConfigurationFromXMLFile(IWin32Window owner)
        {
            if (File.Exists(_fullConfigFilePath))
            {
                XmlDocument configDocument = new XmlDocument();
                configDocument.Load(_fullConfigFilePath);
                _lcid = Convert.ToInt32(configDocument.FirstChild.ChildNodes[0].InnerText);
                _connectToDocumentation = Convert.ToBoolean(configDocument.FirstChild.ChildNodes[1].InnerText);
                _saveSettings = Convert.ToBoolean(configDocument.FirstChild.ChildNodes[2].InnerText);
            }
            else
            { 
                FormOptions dialog = new FormOptions();
                dialog.ShowDialog(owner);
            }
        }

        public static void SaveConfigurationToXMLFile()
        {
            if (File.Exists(_fullConfigFilePath))
                File.Delete(_fullConfigFilePath);

            if (_saveSettings)
            {
                XmlDocument configDocument = new XmlDocument();
                XmlNode firstNode = configDocument.AppendChild(configDocument.CreateElement("Settings"));
                XmlNode lcidNode = firstNode.AppendChild(configDocument.CreateElement("LCID"));
                XmlNode connectNode = firstNode.AppendChild(configDocument.CreateElement("Connect"));
                XmlNode saveNode = firstNode.AppendChild(configDocument.CreateElement("SaveSettings"));
                lcidNode.InnerText = _lcid.ToString();
                connectNode.InnerText = _connectToDocumentation.ToString();
                saveNode.InnerText = _saveSettings.ToString();

                configDocument.Save(_fullConfigFilePath);
            }
        }

        #endregion

        #region Properties

        public static bool ConnectToDocumentation
        {
            get 
            {
                return _connectToDocumentation;
            }
        }
          
        public static bool SaveSettings
        {
            get
            {
                return _saveSettings;
            }
        }

        public static int LCID
        {
            get
            {
                return _lcid;
            }
        }

        public static int DefaultLCID
        {
            get 
            {
                return 1033;
            }
        }

        #endregion

        #region Trigger

        private void radioButtonLanguage1033_CheckedChanged(object sender, EventArgs e)
        {
            _lcid = radioButtonLanguage1033.Checked ? 1033 : 1031;
            Translator.TranslateControls(this, "FormOptions.txt");
        }

        private void radioButtonShowLink_CheckedChanged(object sender, EventArgs e)
        {
            _connectToDocumentation = !radioButtonShowLink.Checked;
        }

        private void checkBoxSaveSettings_CheckedChanged(object sender, EventArgs e)
        {
            _saveSettings = checkBoxSaveSettings.Checked;
        }

        private void buttonDone_Click(object sender, EventArgs e)
        {
            SaveConfigurationToXMLFile();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        #endregion
    }
}
