using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExampleBase
{
    public partial class AreaForm : Form
    {
        public AreaForm()
        {
            InitializeComponent();
        }

        public AreaForm(Control control)
        {
            InitializeComponent();
            Controls.Add(control);
            control.Dock = DockStyle.Fill;
        }

        public static void ShowForm(IWin32Window owner, IExample example)
        {
            AreaForm form = new AreaForm(example.Panel);
            form.Text = example.Caption;
            form.ShowDialog(owner);
            form.Controls.Clear();
            form.Dispose();
        }
    }
}
