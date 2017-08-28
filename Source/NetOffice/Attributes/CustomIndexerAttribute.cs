using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// This indexer is a custom overload from NetOffice
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class CustomIndexerAttribute : System.Attribute
    {
    }
}
