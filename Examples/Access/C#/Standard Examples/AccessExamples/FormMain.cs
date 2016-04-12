using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using ExampleBase;

namespace AccessExamplesCS4
{
    public partial class FormMain : FormBase
    {
        public FormMain()
        {
            InitializeComponent();
            this.Text = "NetOffice Access Examples in C#";
            LoadExamples();
        }

        private void LoadExamples()
        {
            LoadExample(new Example01());
            LoadExample(new Example02());
            LoadExample(new Example03());
            LoadExample(new Example04());
            LoadExample(new Example05());
        }
    }
}
