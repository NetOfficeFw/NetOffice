using System;
using System.Collections.Generic;
using System.Text;

namespace LateBindingApi.Core
{
    /// <summary>
    /// this method is a custom overload from NetOffice
    /// </summary>
    [AttributeUsage(AttributeTargets.All)]
    public sealed class CustomMethodAttribute : System.Attribute
    {
    }
}
