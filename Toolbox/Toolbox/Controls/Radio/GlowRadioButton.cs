using System;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.Controls.Radio
{
    public class GlowRadioButton : System.Windows.Forms.RadioButton
    {
        public GlowRadioButton()
        {
            CheckedChanged += new EventHandler(GlowRadioButton_CheckedChanged);
        }

        private void GlowRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (Checked)
                ForeColor = Color.Blue;
            else
                ForeColor = Color.FromKnownColor(KnownColor.ControlText);
        }
    }
}
