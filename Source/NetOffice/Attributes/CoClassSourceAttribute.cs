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

        --In some rare cases, an interface is a base type for more than 1 CoClass.--
        -- Here is a list of these rare types --
        ----------------------------------------------------------------------------
           - NetOffice.OutlookApi.MAPIFolder (DispatchInterface)


        TODO: Update this comment before release NetOffice 2.0(stable)
    */

    /// <summary>
    /// Indicates an interface is also a well known CoClass
    /// </summary>
    [AttributeUsage(AttributeTargets.Interface, AllowMultiple = false)]
    public class CoClassSourceAttribute : System.Attribute
    {
        /// <summary>
        /// Primary Co Class Type(Interface)
        /// </summary>
        public readonly Type Value;

        /// <summary>
        /// Secondary CoClass Types
        /// </summary>
        public readonly Type[] Values;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="value">primary coclass type</param>
        /// <param name="values">secondary coclass types</param>
        public CoClassSourceAttribute(Type value, params Type[] values)
        {
#if DEBUG
             if(!value.IsInterface)
                throw new ArgumentException("Invalid value.");
#endif
            Value = value;
            Values = values;
        }

        /// <summary>
        /// Returns CoClassSourceAttribute attribute from type or null(Nothing in Visual Basic)
        /// </summary>
        /// <param name="type">target type</param>
        /// <param name="typeId">known type id from target type</param>
        /// <returns>attribute or null(Nothing in Visual Basic)</returns>
        /// <exception cref="ArgumentNullException">type or comProxy is null</exception>
        public static CoClassSourceAttribute TryGet(Type type, Guid typeId)
        {
            if (null == type)
                throw new ArgumentNullException("type");

            CoClassSourceAttribute result = null;

            result = type.GetCustomAttribute<CoClassSourceAttribute>();

            return result;
        }
    }
}
