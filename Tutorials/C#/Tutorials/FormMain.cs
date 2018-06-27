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
    public partial class FormMain : TutorialForm
    {
        public FormMain()
        {
            InitializeComponent();
            Text = "NetOffice Tutorials in C#";
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
            LoadTutorial(new Tutorial14());
            LoadTutorial(new Tutorial15());
            LoadTutorial(new Tutorial16());
            LoadTutorial(new Tutorial17());
            LoadTutorial(new Tutorial18());
        }
    }
}