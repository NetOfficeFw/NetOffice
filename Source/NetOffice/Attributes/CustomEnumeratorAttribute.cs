using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// This enumerator is not supported from the com proxy instance, its a custom service from NetOffice
    /// </summary>
    [AttributeUsage(AttributeTargets.All)]
    public sealed class CustomEnumeratorAttribute : System.Attribute
    {
    }
}
