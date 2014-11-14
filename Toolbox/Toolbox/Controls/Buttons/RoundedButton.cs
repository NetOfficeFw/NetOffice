using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.DeveloperToolbox.Utils.Animation.Round;

namespace NetOffice.DeveloperToolbox.Controls.Buttons
{
    public class RoundedButton : Button
    {
        public RoundedButton()
        {
            RoundedRectangleRegion rndRectRegion = new RoundedRectangleRegion();
            this.Region = rndRectRegion.GetRoundedRect(new RectangleF(this.ClientRectangle.Left, this.ClientRectangle.Top, this.ClientRectangle.Width, this.ClientRectangle.Height), 8);
            this.Resize += new EventHandler(RoundedButton_Resize);
        }

        private void RoundedButton_Resize(object sender, EventArgs e)
        {
            RoundedRectangleRegion rndRectRegion = new RoundedRectangleRegion();
            this.Region = rndRectRegion.GetRoundedRect(new RectangleF(this.ClientRectangle.Left, this.ClientRectangle.Top, this.ClientRectangle.Width, this.ClientRectangle.Height), 8);
        }
    }
}
