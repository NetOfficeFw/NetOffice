using System;

namespace NetOffice
{
    /// <summary>
    /// Indicates original type of entity
    /// </summary>
    public enum EntityType
    {
        /// <summary>
        /// entity is class
        /// </summary>
        IsCoClass = 0,

        /// <summary>
        /// entity is dispatch interface
        /// </summary>
        IsDispatchInterface = 1,

        /// <summary>
        /// entity is interface
        /// </summary>
        IsInterface = 2,

        /// <summary>
        /// entity is module
        /// </summary>
        IsModule = 3,

        /// <summary>
        /// entity is enum
        /// </summary>
        IsEnum = 4,

        /// <summary>
        /// entity is struct
        /// </summary>
        IsStruct = 5,

        /// <summary>
        /// entity is const module
        /// </summary>
        IsConstants = 6
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
