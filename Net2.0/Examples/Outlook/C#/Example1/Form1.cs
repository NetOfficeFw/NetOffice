using System;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;

using LateBindingApi.Core;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums; 

namespace Example1
{
    public partial class Form1 : Form
    {
        Outlook.Application _outlookApplication;

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

            // start outlook and turn off msg boxes
            _outlookApplication = new Outlook.Application();           
 
        }
        
    }
}
