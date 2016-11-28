using System;
using System.Drawing;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Color related helper tools
    /// </summary>
    public class ColorUtils
    {
        #region Fields

        private CommonUtils _owner;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">owner instance</param>
        internal ColorUtils(CommonUtils owner)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            _owner = owner;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public ColorUtils()
        {

        }

        #endregion

        #region Methods

        /// <summary>
        /// Converts a color to its double representation
        /// </summary>
        /// <param name="color">target color to convert</param>
        /// <returns>color argument as double</returns>
        public double ToDouble(System.Drawing.Color color)
        {
            uint returnValue = color.B;
            returnValue = returnValue << 8;
            returnValue += color.G;
            returnValue = returnValue << 8;
            returnValue += color.R;
            return returnValue;
        }

        /// <summary>
        /// Converts a double to its color representation
        /// </summary>
        /// <param name="color">target color to convert</param>
        /// <returns>color argument as System.Drawing.Color</returns>
        public Color ToColor(double color)
        {
            int intColor = Convert.ToInt32(color);
            return System.Drawing.ColorTranslator.FromOle(intColor);
        }

        /// <summary>
        /// Converts a double to its color representation
        /// </summary>
        /// <param name="color">target color to convert</param>
        /// <returns>color argument as System.Drawing.Color</returns>
        public Color ToColor(object color)
        {
            int intColor = Convert.ToInt32(color);
            return System.Drawing.ColorTranslator.FromOle(intColor);
        }

        #endregion
    }
}