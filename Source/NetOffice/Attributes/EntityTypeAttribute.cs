using System;

namespace NetOffice.Attributes
{
    /// <summary>
    /// Indicates original type of entity
    /// </summary>
    public enum EntityType
    {
        /// <summary>
        /// Entity is class
        /// </summary>
        IsCoClass = 0,

        /// <summary>
        /// Entity is dispatch interface
        /// </summary>
        IsDispatchInterface = 1,

        /// <summary>
        /// Entity is interface
        /// </summary>
        IsInterface = 2,

        /// <summary>
        /// Entity is module
        /// </summary>
        IsModule = 3,

        /// <summary>
        /// Entity is enum
        /// </summary>
        IsEnum = 4,

        /// <summary>
        /// Entity is struct
        /// </summary>
        IsStruct = 5,

        /// <summary>
        /// Entity is const module
        /// </summary>
        IsConstants = 6,

        /// <summary>
        /// Entity is native COMImport interface
        /// </summary>
        IsNativeInterface = 7,

        /// <summary>
        /// Entity is wrapper arround a native COMImport interface
        /// That means the instance performs early-bind calls
        /// </summary>
        IsNativeInterfaceCaller = 8
    }

    /// <summary>
    /// Indicates original type of entity in COM Type Library
    /// </summary>
    [AttributeUsage(AttributeTargets.All)]
    public sealed class EntityTypeAttribute : System.Attribute
    {
        /// <summary>
        /// returns type of entity
        /// </summary>
        public readonly EntityType Type;

        /// <summary>
        /// creates instance
        /// </summary>
        /// <param name="type"></param>
        public EntityTypeAttribute(EntityType type)
        {
            this.Type = type;
        }
    }
}
