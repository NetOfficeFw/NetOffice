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
    public partial class FormFinish : Form
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public FormFinish()
        {
            InitializeComponent();
            labelMessage.Text = "Done!";
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="message">Given message as any</param>
        public FormFinish(string message)
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
