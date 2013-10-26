using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.OutlookSecurity
{
    public class SecurityDialog
    {
        internal SecurityDialog(IntPtr handle, string caption, string className, Rect dimension)
        {
            Handle = handle;
            Caption = caption;
            ClassName = className;
            Dimension = dimension;
        }
        public IntPtr Handle { get; internal set; }
        public string Caption { get; internal set; }
        public string ClassName { get; internal set; }
        public Rect Dimension { get; internal set; }
    }
}
