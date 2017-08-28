using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Indicates object typed result is a always a COM reference or null(Nothing in Visual Basic)
    /// </summary>
    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Property)]
    public class ProxyResultAttribute : System.Attribute
    {

    }
}
