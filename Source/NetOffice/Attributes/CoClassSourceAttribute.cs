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

        (In some rare cases, an interface is a base type for more than 1 CoClass.)

        TODO: Update this comment before release NetOffice 2.0(stable)
    */

    /// <summary>
    /// Indicates an interface is also a well known CoClass
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
#if DEBUG
             if(!value.IsInterface)
                throw new ArgumentException("Invalid value.");
#endif
            Value = value;
        }

        /// <summary>
        /// Returns CoClassSourceAttribute attribute from type or null
        /// </summary>
        /// <param name="type">target type</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException">type is null</exception>
        /// <exception cref="ArgumentException">type does not have a typeid attribute</exception>
        /// <exception cref="NotSupportedException">type is source for more than 1 coclass. this is currently unsupported</exception>
        internal static CoClassSourceAttribute TryGet(Type type)
        {
            var typeId = type.GetCustomAttribute<TypeIdAttribute>();
            if (null == typeId)
                throw new ArgumentException("Invalid type.");
            return TryGet(type, typeId.Value);
        }

        /// <summary>
        /// Returns CoClassSourceAttribute attribute from type or null(Nothing in Visual Basic)
        /// </summary>
        /// <param name="type">target type</param>
        /// <param name="typeId">known type id from target type</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException">type is null</exception>
        /// <exception cref="NotSupportedException">type is source for more than 1 coclass. this is currently unsupported</exception>
        public static CoClassSourceAttribute TryGet(Type type, Guid typeId)
        {
            if (null == type)
                throw new ArgumentNullException("type");

            var coClasses = type.GetCustomAttributes<CoClassSourceAttribute>();
            if (coClasses.Length > 1)
            {
                throw new NotSupportedException();
            }
            else if (coClasses.Length == 1)
            {
                return coClasses.First();
            }
            else
            {
                return null;
            }
        }
    }
}
