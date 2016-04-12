using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using TutorialsBase;

namespace TutorialsCS4
{
    public partial class FormMain : FormBase
    {
        public FormMain()
        {
            InitializeComponent();
            this.Text = "NetOffice Tutorials in C#";
            LoadTutorials();
        }

        private void LoadTutorials()
        {
            LoadTutorial(new Tutorial01());
            LoadTutorial(new Tutorial02());
            LoadTutorial(new Tutorial03());
            LoadTutorial(new Tutorial04());
            LoadTutorial(new Tutorial05());
            LoadTutorial(new Tutorial06());
            LoadTutorial(new Tutorial07());
            LoadTutorial(new Tutorial08());
            LoadTutorial(new Tutorial09());
            LoadTutorial(new Tutorial10());
            LoadTutorial(new Tutorial11());
            LoadTutorial(new Tutorial12());
            LoadTutorial(new Tutorial13());
            NavigateToTutorial(0);
        }
    }
}
