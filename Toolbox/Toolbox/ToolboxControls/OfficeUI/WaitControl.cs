using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.OfficeUI
{
    /// <summary>
    /// A wait panel while office application is creating and fetching the ui object model
    /// </summary>
    public partial class WaitControl : UserControl
    {
        #region Fields

        private int _currentLanguageID;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="currentLanguageID">current user language id</param>
        public WaitControl(int currentLanguageID)
        {
            InitializeComponent();
            ResizeControls();
            _currentLanguageID = currentLanguageID;
            Translation.Translator.TranslateControls(this, "ToolboxControls.OfficeUI.WaitControlTable.txt", _currentLanguageID);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Current user language id
        /// </summary>
        public int CurrentLanguageID
        {
            get
            {
                return _currentLanguageID;
            }
            set
            {
                _currentLanguageID = value;
                Translation.Translator.TranslateControls(this, "ToolboxControls.OfficeUI.WaitControlTable.txt", _currentLanguageID);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Bring control in front
        /// </summary>
        public new void Show()
        {
            labelWaitMessage.Visible = false;
            base.Show();
            ResizeControls();
        }

        /// <summary>
        /// Report current action
        /// </summary>
        /// <param name="message">current action to display for the user</param>
        public void ReportProgress(string message)
        {
            labelWaitMessage.Text = message;
            labelWaitMessage.Visible = true;
            labelWaitMessage.Refresh();
        }

        private void ResizeControls()
        {
            pictureBoxLogo.Left = (this.Width / 2) - (pictureBoxLogo.Width / 2);
            pictureBoxLogo.Top = (this.Height / 2) - (pictureBoxLogo.Height / 2);

            labelHeader.Left = (this.Width / 2) - (labelHeader.Width / 2);
            labelHeader.Top = pictureBoxLogo.Top - 50;

            labelWaitMessage.Left = (this.Width / 2) - (labelHeader.Width / 2);
            labelWaitMessage.Top = pictureBoxLogo.Top + pictureBoxLogo.Height + 50;
        }

        #endregion

        #region Trigger

        private void WaitControl_Resize(object sender, EventArgs e)
        {
            ResizeControls();
        }

        #endregion
    }
}
