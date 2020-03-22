using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Determine type of sink interface method argument
    /// </summary>
    public enum SinkArgumentType
    {
        /// <summary>
        /// System.Int16
        /// </summary>
        Int16 = 0,

        /// <summary>
        /// System.Int32
        /// </summary>
        Int32 = 1,

        /// <summary>
        /// System.Single
        /// </summary>
        Single = 2,

        /// <summary>
        /// System.Double
        /// </summary>
        Double = 3,

        /// <summary>
        /// System.String
        /// </summary>
        String = 4,

        /// <summary>
        /// System.Boolean
        /// </summary>
        Bool = 5,

        /// <summary>
        /// System.Enum
        /// </summary>
        Enum = 6,

        /// <summary>
        /// Known Proxy Type
        /// </summary>
        KnownProxy = 7,

        /// <summary>
        /// Unknown Proxy Type
        /// </summary>
        UnknownProxy = 8
    }

    /// <summary>
    /// Sink Interface Argument Information
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
    public class SinkArgumentAttribute : System.Attribute
    {
        /// <summary>
        /// Argument Name
        /// </summary>
        public readonly string Name;

        /// <summary>
        /// Argument Type
        /// </summary>
        public readonly SinkArgumentType Type;

        /// <summary>
        /// Target Type - if its
        /// </summary>
        public readonly Type Convert;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="name">argument name</param>
        /// <param name="type">argument type</param>
        public SinkArgumentAttribute(string name, SinkArgumentType type)
        {
            Name = name;
            Type = type;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="name">argument name</param>
        /// <param name="convert">argument convert to</param>
        public SinkArgumentAttribute(string name, Type convert)
        {
            Name = name;
            Type = SinkArgumentType.KnownProxy;
            Convert = convert;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="name">argument name</param>
        /// <param name="type">argument type</param>
        /// <param name="convert">argument convert to</param>
        public SinkArgumentAttribute(string name, SinkArgumentType type, Type convert)
        {
            Name = name;
            Type = type;
            Convert = convert;
        }
    }
}
