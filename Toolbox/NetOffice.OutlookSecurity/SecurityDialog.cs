using System;

namespace NetOffice.OutlookSecurity
{
    /// <summary>
    /// Matching info for outlook security dialog
    /// </summary>
    public class SecurityDialog
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="handle">handle from the dialog</param>
        /// <param name="caption">caption from the dialog</param>
        /// <param name="className">class name from the dialog</param>
        /// <param name="dimension">postion/size from the dialog</param>
        internal SecurityDialog(IntPtr handle, string caption, string className, Rect dimension)
        {
            Handle = handle;
            Caption = caption;
            ClassName = className;
            Dimension = dimension;
        }

        /// <summary>
        /// Dialog handle
        /// </summary>
        public IntPtr Handle { get; internal set; }

        /// <summary>
        /// Dialog caption
        /// </summary>
        public string Caption { get; internal set; }

        /// <summary>
        /// Dialog class name
        /// </summary>
        public string ClassName { get; internal set; }


        /// <summary>
        /// Dialog position/size
        /// </summary>
        public Rect Dimension { get; internal set; }
    }
}
