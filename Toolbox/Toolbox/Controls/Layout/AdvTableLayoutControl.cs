using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Controls.Layout
{
    /// <summary>
    /// Standard TableLayoutPanel but its flicker-free
    /// </summary>
    public partial class AdvTableLayoutControl : TableLayoutPanel
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public AdvTableLayoutControl()
        {
            InitializeComponent();
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);
        }
    }
}
