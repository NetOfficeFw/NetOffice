using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Dialogs
{
    /// <summary>
    /// Color settings in default dialogs
    /// </summary>
    public class DialogLayoutSettings
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public DialogLayoutSettings()
        {
            BackColor = Color.LightSteelBlue;
            BackAlternateColor = Color.Orange;
            BackHeaderColor = Color.White;
            ForeColor = Color.Black;
            ForeAlternateColor = Color.Blue;
        }

        #endregion

        #region Properties

        /// <summary>
        /// BackColor in default dialogs. Default: Color.LightSteelBlue
        /// </summary>
        public Color BackColor { get; set; }

        /// <summary>
        /// Alternate BackColor in default dialogs. Default: Color.Orange
        /// </summary>
        public Color BackAlternateColor { get; set; }

        /// <summary>
        /// Header element BackColor in default dialogs. Default: Color.White
        /// </summary>
        public Color BackHeaderColor { get; set; }

        /// <summary>
        /// ForeColor in default dialogs. Default: Color.Black
        /// </summary>
        public Color ForeColor { get; set; }

        /// <summary>
        /// Alternate ForeColor default dialogs. Default: Color.Blue
        /// </summary>
        public Color ForeAlternateColor { get; set; }

        #endregion
    }
}
