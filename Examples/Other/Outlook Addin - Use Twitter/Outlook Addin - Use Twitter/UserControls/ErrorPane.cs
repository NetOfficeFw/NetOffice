using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Sample.Addin
{
    public partial class ErrorPane : UserControl
    {
        public ErrorPane()
        {
            InitializeComponent();
        }

        public void ShowError(Exception exception)
        {
            pictureBox2.Visible = true;
            labelErrorMessage.Text = exception.Message;
            labelErrorMessage.Visible = true;
            buttonErrorDetails.Visible = true;
        }

        public void ClearError()
        {
            pictureBox2.Visible = false;
            labelErrorMessage.Text = string.Empty;
            labelErrorMessage.Visible = false;
            buttonErrorDetails.Visible = false;
        }
    }
}
