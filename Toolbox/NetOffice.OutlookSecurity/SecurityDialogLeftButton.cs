using System;

namespace NetOffice.OutlookSecurity
{
    /// <summary>
    /// Represents the bottom left button in the outlook security dialog
    /// </summary>
    public class SecurityDialogLeftButton
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="handle">button handle</param>
        /// <param name="caption">button caption</param>
        /// <param name="dimension">button location/size</param>
        internal SecurityDialogLeftButton(IntPtr handle, string caption, Rect dimension)
        {
            Handle = handle;
            Caption = caption;
            Dimension = dimension;
        }

        /// <summary>
        /// Button handle
        /// </summary>
        public IntPtr Handle { get; internal set; }

        /// <summary>
        /// Button caption
        /// </summary>
        public string Caption { get; internal set; }

        /// <summary>
        /// Button location/size
        /// </summary>
        public Rect Dimension { get; internal set; }
    }
}
