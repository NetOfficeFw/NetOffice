using System;
using System.Reflection;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.OfficeApi.Tools.Dialogs
{
    /// <summary>
    /// NetOffice Tools Base Dialog
    /// </summary>
    public partial class ToolsDialog : Form
    {        
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public ToolsDialog()
        {
            InitializeComponent();
        }

        #endregion

        #region Methods

        /// <summary>
        /// Do layout settings in form instance
        /// </summary>
        /// <param name="layout">layout settings</param>
        protected internal virtual void DoLayout(DialogLayoutSettings layout)
        {
            BackColor = layout.BackColor;
            ForeColor = layout.ForeColor;
            DoLayoutControls(this.Controls, layout);
        }

        /// <summary>
        /// Do localization in form instance
        /// </summary>
        /// <param name="localization">localized values</param>
        protected internal virtual void DoLocalization(DialogLocalization localization)
        { 
            
        }

        /// <summary>
        /// Get name, ressource path root schema for all NetOffice default dialogs
        /// </summary>
        /// <returns>enumerator to iterate the schema</returns>
        internal static IEnumerable<string> CreateDialogSchema()
        {
            List<string> result = new List<string>();
            Type type = typeof(ToolsDialog);

            result.Add("DiagnosticsDialog");
            result.Add("ErrorDialog");
            result.Add("AboutDialog");
            result.Add("RichTextDialog");

            return result;
        }

        /// <summary>
        /// Shows an error in a standard message box
        /// </summary>
        /// <param name="error">error description</param>
        protected void ShowSingleException(Exception error)
        {
            if(null != error)
                MessageBox.Show(this, error.ToString(), Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void DoLayoutControls(Control.ControlCollection controls, DialogLayoutSettings layout)
        {
            foreach (Control control in controls)
            {
                if (control.Name == "panelHeader")
                    control.BackColor = layout.BackHeaderColor;

                Button buttonControl = control as Button;
                if (null != buttonControl)
                {
                    buttonControl.ForeColor = layout.ForeAlternateColor;
                    buttonControl.FlatAppearance.BorderColor = layout.BackHeaderColor;
                }
                else
                {
                    control.ForeColor = layout.ForeColor;
                }

                if (control.Controls.Count > 0)
                    DoLayoutControls(control.Controls, layout);
            }
        }

        #endregion
    }
}
