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
    /// User selection completed event handler
    /// </summary>
    /// <param name="officeAppName">target application name</param>
    public delegate void SelectOfficeEventHandler(string officeAppName);

    /// <summary>
    /// Shows supported office application to create an analyze one of them
    /// </summary>
    [RessourceTable("ToolboxControls.OfficeUI.SelectOfficeAppControlStrings.txt")]
    public partial class SelectOfficeAppControl : UserControl
    {
        #region Fields

        private SelectOfficeEventHandler _eventHandler;
    
        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public SelectOfficeAppControl( )
        {
            InitializeComponent();
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="handler">close handler</param>
        public SelectOfficeAppControl(SelectOfficeEventHandler handler)
        {
            InitializeComponent();
            _eventHandler = handler;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Name of selected application
        /// </summary>
        public string SelectedApplication
        {
            get 
            {
                return listView1.SelectedItems[0].Text;
            }
        }

        /// <summary>
        /// Indicates yes want to proceed or abort
        /// </summary>
        public DialogResult Result { get; set; }

        #endregion

        #region Trigger

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
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        #endregion
    }
}
