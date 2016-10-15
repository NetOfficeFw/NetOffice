using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Indicates a static method is to call from RegAddin while unregister addin
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false, Inherited = false)]
    public sealed class ComUnregisterCallAttribute : System.Attribute
    {
    }
}
