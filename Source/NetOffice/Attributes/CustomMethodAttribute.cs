using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// This method is a custom overload from NetOffice
    /// </summary>
    [AttributeUsage(AttributeTargets.Method)]
    public sealed class CustomMethodAttribute : System.Attribute
    {
    }
}
