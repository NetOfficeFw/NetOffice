using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.OutlookSecurity
{
    public class SecurityDialogLeftButton
    {
        internal SecurityDialogLeftButton(IntPtr handle, string caption, Rect dimension)
        {
            Handle = handle;
            Caption = caption;
            Dimension = dimension;
        }
        public IntPtr Handle { get; internal set; }
        public string Caption { get; internal set; }
        public Rect Dimension { get; internal set; }
    }
}
