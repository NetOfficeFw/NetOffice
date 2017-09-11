using System;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Indicates a method or property is overriden from NetOffice Core
    /// </summary>
    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Property)]
    public class CoreOverriddenAttribute : System.Attribute
    {

    }
}
