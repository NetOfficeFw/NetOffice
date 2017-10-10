using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OutlookApi.Tools.Contribution.Security
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

        /// <summary>
        /// Flag to keep information a TimeoutException was thrown for the dialog
        /// </summary>
        internal bool ExceptionThrown { get; set; }

        /// <summary>
        /// Flag to keep information checkbox is clicked for the dialog
        /// </summary>
        internal bool CheckBoxPassed { get; set; }

        /// <summary>
        /// Returns a System.String that represents the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            return String.Format("SecurityDialog {0}", Handle);
        }
    }
}
