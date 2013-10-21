using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;

namespace TutorialsBase
{
    public class TutorialForm : Form
    {
        private System.ComponentModel.IContainer components = null;
      
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        internal System.ComponentModel.IContainer Components
        {
            get
            {
                return this.components;
            }
        }
    }
}
