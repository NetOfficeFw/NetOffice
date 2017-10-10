using System;

namespace NetOffice.OutlookApi.Tools.Contribution.Security
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

        /// <summary>
        /// Returns a System.String that represents the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            return String.Format("SecurityDialogLeftButton {0}", Handle);
        }
    }
}