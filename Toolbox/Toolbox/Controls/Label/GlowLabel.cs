using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using System.Text;

namespace NetOffice.DeveloperToolbox.Controls.Label
{
    /// <summary>
    /// Standard label with a border rectangle if mouse inside
    /// </summary>
    public class GlowLabel : System.Windows.Forms.Label
    {
        #region Fields

        private bool _mouseInside;
        private Color _glowColor = Color.FromArgb(209, 253, 205);
        private Pen _pen;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public GlowLabel()
        {
            _pen = new Pen(_glowColor, 1);
            this.MouseEnter += new EventHandler(GlowLabel_MouseEnter);
            this.MouseLeave += new EventHandler(GlowLabel_MouseLeave);
        }

        #endregion

        #region Overrides

        protected override void OnPaint(PaintEventArgs e)
        {
            if (_mouseInside)
            {
                Rectangle rect = ClientRectangle;
                e.Graphics.DrawRectangle(_pen, rect.X, rect.Y, rect.Width - 1, rect.Height - 1);
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

        #endregion

        #region Trigger

        private void GlowLabel_MouseLeave(object sender, EventArgs e)
        {
            try
            {
                _mouseInside = false;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }
        }

        private void GlowLabel_MouseEnter(object sender, EventArgs e)
        {
            try
            {
                _mouseInside = true;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }
        }

        #endregion
    }
}
