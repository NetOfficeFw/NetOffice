using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace TutorialsBase
{
    /// <summary>
    /// Application Tutorial finished dialog
    /// </summary>
    public partial class FinishForm : Form
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public FinishForm()
        {
            InitializeComponent();
            labelMessage.Text = "Done!";
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">Given message as any</param>
        public FinishForm(string message)
        {
            InitializeComponent();

            if (null == message)
                message = "Done!";
           
            labelMessage.Text = message;
        }

        #endregion

        #region Trigger

        private void buttonClose_Click(object sender, EventArgs e)
        {
            try
            {
                this.Close();
            }
            catch
            {
                MessageBox.Show(this, "An error occured.", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion
    }
}
