using System;
using System.Windows.Forms;
using ExampleBase;

namespace OutlookExamplesCS4
{
    public partial class FormMain : ExampleForm
    {
        public FormMain()
        {
            InitializeComponent();
            this.Text = "NetOffice Outlook Examples in C#";
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
