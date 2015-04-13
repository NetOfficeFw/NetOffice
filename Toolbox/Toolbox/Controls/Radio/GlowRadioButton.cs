using System;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.Controls.Radio
{
    /// <summary>
    /// Standard radio button which is blue if checked
    /// </summary>
    public class GlowRadioButton : System.Windows.Forms.RadioButton
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public GlowRadioButton()
        {
            CheckedChanged += new EventHandler(GlowRadioButton_CheckedChanged);
        }

        #endregion

        #region Trigger

        private void GlowRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (Checked)
                ForeColor = Color.Blue;
            else
                ForeColor = Color.FromKnownColor(KnownColor.ControlText);
        }

        #endregion
    }
}
