using System;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;

using LateBindingApi.Core;
using PowerPoint = NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums; 

namespace Example1
{
    public partial class Form1 : Form
    {
        PowerPoint.Application _powerApplication;

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
            _powerApplication = new PowerPoint.Application();
     
        }
    }
}
