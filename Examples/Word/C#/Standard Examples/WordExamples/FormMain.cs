using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using ExampleBase;

namespace WordExamplesCS4
{
    public partial class FormMain : ExampleForm
    {
        public FormMain()
        {
            InitializeComponent();

            this.Text = "NetOffice Word Examples in C#";
            LoadExamples();

        }

        private void LoadExamples()
        {
            LoadExample(new Example01());
            LoadExample(new Example02());
            LoadExample(new Example03());
            LoadExample(new Example04());
            LoadExample(new Example05());
            LoadExample(new Example06());
            LoadExample(new Example07());
        }
    }
}
