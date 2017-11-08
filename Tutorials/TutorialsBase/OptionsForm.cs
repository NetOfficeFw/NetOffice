using System;
using System.IO;
using System.Xml;
using System.Windows.Forms;

namespace TutorialsBase
{
    /// <summary>
    /// Application user options form
    /// </summary>
    partial class OptionsForm : Form
    {
        #region Fields

        private static string _fullConfigFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) , "NetOfficeTutorialsCS4.xml");
        private static bool _connectToDocumentation = false;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public OptionsForm()
        {
            InitializeComponent();
            if (!_connectToDocumentation)
                radioButtonShowLink.Checked = true;
        }

        #endregion

        #region Properties

        /// <summary>
        /// User allows to connect the NetOffice tutorial pages
        /// </summary>
        public static bool ConnectToDocumentation
        {
            get 
            {
                return _connectToDocumentation;
            }
        }
          

        #endregion

        #region Trigger

        private void radioButtonShowLink_CheckedChanged(object sender, EventArgs e)
        {
            _connectToDocumentation = !radioButtonShowLink.Checked;
        }

        private void buttonDone_Click(object sender, EventArgs e)
        { 
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        #endregion
    }
}
