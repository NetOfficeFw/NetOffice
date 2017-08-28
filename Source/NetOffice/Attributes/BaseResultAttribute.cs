using System;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Indicates the property or method return type is a base type.
    /// </summary>
    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Property)]
    public class BaseResultAttribute : System.Attribute
    {

    }
}
