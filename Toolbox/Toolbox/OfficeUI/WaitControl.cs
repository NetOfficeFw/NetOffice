using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox
{
    public partial class WaitControl : UserControl
    {
        int _currentLanguageID;

        public WaitControl(int currentLanguageID)
        {
            InitializeComponent();
            ResizeControls();
            _currentLanguageID = currentLanguageID;
            Translator.TranslateControls(this, "OfficeUI.WaitControlTable.txt", _currentLanguageID);
        }

        public int CurrentLanguageID
        {
            get
            {
                return _currentLanguageID;
            }
            set
            {
                _currentLanguageID = value;
                Translator.TranslateControls(this, "OfficeUI.WaitControlTable.txt", _currentLanguageID);
            }
        }

        private void WaitControl_Resize(object sender, EventArgs e)
        {
            ResizeControls();
        }

        public new void Show()
        {          
            labelWaitMessage.Visible = false;
            base.Show();
            ResizeControls();
        }

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
    }
}
