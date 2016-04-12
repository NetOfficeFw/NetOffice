using System;
using System.Drawing;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Image related utils
    /// </summary>
    public class ImageUtils
    {
        #region ConvertImage

        private class ConvertImage : System.Windows.Forms.AxHost
        {
            private ConvertImage() : base(null)
            {
            }

            public static stdole.IPictureDisp Convert(Image image)
            {
                return System.Windows.Forms.AxHost.GetIPictureDispFromPicture(image) as stdole.IPictureDisp;
            }
        }

        #endregion

        #region Fields

        private CommonUtils _owner;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">owner instance</param>
        internal ImageUtils(CommonUtils owner)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            _owner = owner;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Converts an image to IPictureDisp
        /// </summary>
        /// <param name="image">target image to convert</param>
        /// <returns>IPictureDisp instance</returns>
        public stdole.IPictureDisp ToPicture(Image image)
        {
            if (null == image)
                throw new ArgumentNullException("image");
            return ConvertImage.Convert(image);
        }

        /// <summary>
        /// Copy given image and converts extracted mask to IPictureDisp
        /// </summary>
        /// <param name="image">target image</param>
        /// <returns>IPictureDisp instance</returns>
        public stdole.IPictureDisp ToPictureMask(Image image)
        {
            if (null == image)
                throw new ArgumentNullException("image");
            
            return ToPicture(ToMask(image));
        }

        /// <summary>
        /// Copy given image and returns extracted black/white mask
        /// </summary>
        /// <param name="image">target image</param>
        /// <returns>image mask copy</returns>
        public Image ToMask(Image image)
        {
            if (null == image)
                throw new ArgumentNullException("image");

            Bitmap target = new Bitmap(image);

            for (int y = 0; y < target.Height; y++)
            {
                for (int x = 0; x < target.Width; x++)
                {
                    Color color = target.GetPixel(x, y);
                    Color mask = Color.Empty;

                    if (color.B == 0 && color.R == 0 && color.G == 0)
                        mask = Color.White;
                    else
                        mask = Color.Black;

                    target.SetPixel(x, y, Color.FromArgb(color.A, mask.R, mask.G, mask.B));
                }
            }

            return target;
        }

        #endregion
    }
}
