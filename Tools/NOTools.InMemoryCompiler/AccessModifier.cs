using System;

namespace NOTools.InMemoryCompiler
{
    /// <summary>
    /// Access modifier for classes, methods and properties
    /// </summary>
    public enum AccessModifier
    {
        /// <summary>
        /// Target is public
        /// </summary>
        Public = 0,
        
        /// <summary>
        /// Target is internal
        /// </summary>
        Internal = 1,
        
        /// <summary>
        /// Target is private
        /// </summary>
        Private = 2
    }
}
