using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Utils.Animation.Panel
{
    internal class ImageShape : Shape
    {
        #region Fields

        private float _opacity;
        private ImageAttributes _attributes;

        #endregion

        #region Properties

        public Image Image { get; set; }

        public float Opacity
        {
            get 
            {
                return _opacity;
            }
            set
            {
                if (value < 0.1f || value > 1.0f)
                    throw new ArgumentException();
                _opacity = value;
                if (_opacity < 1.0f)
                {
                    ColorMatrix matrix = new ColorMatrix();
                    matrix.Matrix33 = _opacity;
                    _attributes = new ImageAttributes();
                    _attributes.SetColorMatrix(matrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);
                }
                else
                {
                    _attributes = null;
                }
            }
        }

        #endregion

        #region Overrides

        public override void RenderObject(Graphics graphic)
        {
            if (null != Image)
            {
                if (null != _attributes)
                {
                    graphic.DrawImage(Image, new Rectangle(0, 0, Image.Width, Image.Height), 0, 0, Image.Width, Image.Height, GraphicsUnit.Pixel, _attributes);
                }
                else
                {
                    graphic.DrawImage(Image, new Rectangle(Location.X, Location.Y, Size.Width, Size.Height));
                }            
            }
        }

        #endregion
    }
}
