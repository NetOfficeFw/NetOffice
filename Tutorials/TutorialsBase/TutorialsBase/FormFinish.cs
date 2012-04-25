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
    public partial class FormFinish : Form
    {
        public FormFinish()
        {
            InitializeComponent();
            labelMessage.Text = "Done!";
        }

        public FormFinish(string message)
        {
            InitializeComponent();

            if (null == message)
                message = "Done!";
           
            labelMessage.Text = message;
        }

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
    }
}
