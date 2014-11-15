using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using System.Text;

namespace NetOffice.DeveloperToolbox.Controls.Label
{
    public class GlowLabel : System.Windows.Forms.Label
    {
        private bool _mouseInside;
        private Color _glowColor = Color.FromArgb(209, 253, 205);
        private Pen _pen;

        public GlowLabel()
        {
            _pen = new Pen(_glowColor, 1);
            this.MouseEnter += new EventHandler(GlowLabel_MouseEnter);
            this.MouseLeave += new EventHandler(GlowLabel_MouseLeave);
        }

        private void GlowLabel_MouseLeave(object sender, EventArgs e)
        {
            _mouseInside = false;
        }

        private void GlowLabel_MouseEnter(object sender, EventArgs e)
        {
            _mouseInside = true;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            if (_mouseInside)
            {
                Rectangle rect= ClientRectangle;
                e.Graphics.DrawRectangle(_pen, rect.X, rect.Y, rect.Width-1, rect.Height-1);
            }
            base.OnPaint(e);
        }
        protected override void Dispose(bool disposing)
        {
            if (null != _pen)
            { 
                _pen.Dispose();
                _pen = null;
            }
            base.Dispose(disposing);
        }
    }
}
