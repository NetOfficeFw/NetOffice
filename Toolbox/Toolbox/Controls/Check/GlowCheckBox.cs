using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Controls.Check
{    
    public class GlowCheckBox : System.Windows.Forms.CheckBox
    {
        private Pen _selectedPen;

        public GlowCheckBox()
        {
            _selectedPen = new Pen(Color.Blue, 1);
        }

        protected override void OnPaint(PaintEventArgs pevent)
        {
            base.OnPaint(pevent);

            int offset = 2;
            SizeF stringMeasure = pevent.Graphics.MeasureString(Text, Font);

            int leftOffset = offset + Padding.Left;
            int topOffset = (int)(ClientRectangle.Height - stringMeasure.Height) / 2;
            if (topOffset < 0)
                topOffset = offset + Padding.Top;
            else
                topOffset += Padding.Top;
            
            if (Checked)
                pevent.Graphics.DrawRectangle(_selectedPen, 0, topOffset + 4, 10, 10);
        }

        protected override void Dispose(bool disposing)
        {
            if (null != _selectedPen)
            {
                _selectedPen.Dispose();
                _selectedPen = null;
            }
            base.Dispose(disposing);
        }
    }
}
