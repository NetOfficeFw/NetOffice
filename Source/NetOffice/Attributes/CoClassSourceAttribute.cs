using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /*
        This is to bypass the issue that the type id of a com proxy is always the
        interface id, never the coclass id, but we need to create the highest possible wrapper type
        to make sure the caller can cast the result in any type he want.
        
        If the NetOffice Type Resolver find an interface type with CoClassSourceAttribute then
        he return the CoClass type instead.

        This is first used -today- by NetOffice 2.0 pre-alpha(experimental) in order to change the name-based wrapper resolving
        to an id-based resolving(possibly faster). If it doesnt works, this attribute is unused in the future and we go back
        to name-based resolving.

        TODO: Update this comment before release NetOffice 2.0(stable)
    */

    /// <summary>
    /// Indicates an interface is a well known CoClass implementation
    /// </summary>
    [AttributeUsage(AttributeTargets.Interface, AllowMultiple = true)]
    public class CoClassSourceAttribute : System.Attribute
    {
        /// <summary>
        /// Co Class Type(Interface)
        /// </summary>
        public readonly Type Value;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="value">coclass type</param>
        public CoClassSourceAttribute(Type value)
        {
            Value = value;
        }
    }
}
