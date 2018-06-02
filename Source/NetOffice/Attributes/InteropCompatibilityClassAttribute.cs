using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Class is an interop assembly pedant to be more familar with existing PIA codebases
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class InteropCompatibilityClassAttribute : System.Attribute
    {
    }
}
