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
    public delegate void SelectOfficeEventHandler(string officeAppName);

    public partial class SelectOfficeAppControl : UserControl
    {
        SelectOfficeEventHandler _eventHandler;
        int _currentLanguageID;

        public SelectOfficeAppControl(int currentLanguageID, SelectOfficeEventHandler handler)
        {
            InitializeComponent();
            _eventHandler = handler;
            _currentLanguageID = currentLanguageID;
            Translator.TranslateControls(this, "OfficeUI.SelectOfficeAppControlTable.txt", _currentLanguageID);
        }

        public string SelectedApplication
        {
            get 
            {
                return listView1.SelectedItems[0].Text;
            }
        }

        public DialogResult Result { get; set; }
        
        private void buttonClose2_Click(object sender, EventArgs e)
        {
            Result = DialogResult.Cancel;
            this.Hide();
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            Result = DialogResult.Cancel;
            this.Hide();
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            buttonSelect.Enabled  = (listView1.SelectedIndices.Count > 0);            
        }

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if ((listView1.SelectedIndices.Count > 0))
                buttonSelect_Click(this, new EventArgs());
        }

        private void buttonSelect_Click(object sender, EventArgs e)
        {
            try
            {
                Result = DialogResult.OK;
                this.Hide();
                _eventHandler(SelectedApplication);
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
        }
    }
}
