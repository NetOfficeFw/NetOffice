using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// Runtime Entity Description
    /// </summary>
    public class DynamicObjectEntity
    {
        /// <summary>
        /// Particular Entity Kind
        /// </summary>
        public enum EntityKind
        {
            /// <summary>
            /// Method
            /// </summary>
            Method = 0,

            /// <summary>
            /// ReadOnly Property
            /// </summary>
            PropertyReadonly = 2,

            /// <summary>
            /// Property with Write-Access
            /// </summary>
            PropertyWritable = 3
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="name">entity name</param>
        /// <param name="kind">entity kind</param>
        public DynamicObjectEntity(string name, EntityKind kind)
        {
            Name = name;
            Kind = kind;
        }

        /// <summary>
        /// Entity Name
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Entity Kind
        /// </summary>
        public EntityKind Kind { get; internal set; }

        /// <summary>
        /// Returns a System.String that represents the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            return Name;
        }
    }

}
