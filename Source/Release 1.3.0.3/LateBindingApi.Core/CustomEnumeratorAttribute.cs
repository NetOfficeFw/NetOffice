using System;
using System.Collections.Generic;
using System.Text;

namespace LateBindingApi.Core
{
    /// <summary>
    /// this enumerator is not supported by the instance, its a custom service by NetOffice
    /// </summary>
    [AttributeUsage(AttributeTargets.All)]
    public sealed class CustomEnumeratorAttribute : System.Attribute
    {
    }
}
