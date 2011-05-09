using System;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;

using LateBindingApi.Core;
using Access = NetOffice.AccessApi;
using NetOffice.AccessApi.Enums; 

namespace Example1
{
    public partial class Form1 : Form
    {
        Access.Application _accessApplication;

        public Form1()
        {
            InitializeComponent();

            /*
             * Initialize Api COMObject & COMVariant Support
             */
            LateBindingApi.Core.Factory.Initialize();
        }
  
        private void button1_Click(object sender, EventArgs e)
        {
            // start word and turn off msg boxes
            _accessApplication = new Access.Application();
     
        }
    }
}
