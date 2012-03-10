using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace Host
{
    partial class ExceptionDialog : Form
    {
        public ExceptionDialog(Exception exception)
        {
            InitializeComponent();
            string result = "";
            Exception ex = exception;
            while (ex != null)
            {
                string type = ex.GetType().Name;
                string message = ex.Message;
                string target = "<Empty>";
                if (null != ex.TargetSite)
                    target = ex.TargetSite.ToString();
                string trace = "<Empty>";
                if (null != ex.StackTrace)
                    trace = ex.StackTrace.ToString();

                result += "Type:" + type + Environment.NewLine;
                result += "Message:" + message + Environment.NewLine;
                result += "Target:" + target + Environment.NewLine;
                result += "Stack:" + trace + Environment.NewLine;

                result += Environment.NewLine;
                ex = ex.InnerException;
            }
            textBox1.Text =  result;            
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
