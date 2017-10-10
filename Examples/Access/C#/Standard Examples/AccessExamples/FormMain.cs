using System;
using ExampleBase;

namespace AccessExamplesCS4
{
    public partial class FormMain : ExampleForm
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