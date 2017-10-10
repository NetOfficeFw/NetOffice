using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OutlookApi.Tools.Contribution.Security
{
    /// <summary>
    /// Represents the 'Allow' checkbox from outlook security dialog
    /// </summary>
    public class SecurityDialogCheckBox
    {
        /// <summary>
        /// Creates an instance of the class    
        /// </summary>
        /// <param name="handle">control handle</param>
        /// <param name="caption">control text</param>
        /// <param name="dimension">control location/size</param>
        internal SecurityDialogCheckBox(IntPtr handle, string caption, Rect dimension)
        {
            Handle = handle;
            Caption = caption;
            Dimension = dimension;
        }

        /// <summary>
        /// Control handle
        /// </summary>
        public IntPtr Handle { get; internal set; }

        /// <summary>
        /// Control caption
        /// </summary>
        public string Caption { get; internal set; }

        /// <summary>
        /// Control dimension
        /// </summary>
        public Rect Dimension { get; internal set; }

        /// <summary>
        /// Returns a System.String that represents the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            return String.Format("SecurityDialogCheckBox {0}", Handle);
        }
    }
}
