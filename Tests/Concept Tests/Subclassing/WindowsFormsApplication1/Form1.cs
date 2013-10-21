using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        private Form2 _form2;
        private SubClassingWindow _subClass;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (null != _form2)
                return;
            _form2 = new Form2();
            _subClass = new SubClassingWindow(_form2.UserControlTest.Handle, LogMessage,
                new WndMessage[] { WndMessage.WM_MOUSEMOVE, WndMessage.WM_LBUTTONDOWN });
            _form2.Show();
        }

        private void LogMessage(IntPtr handle, WndMessage message)
        {
            string recievedMessage = message.ToString();
            textBox1.Text = recievedMessage + Environment.NewLine + textBox1.Text;
        }
    }
}
