using System;

namespace NetOffice.ComTypes
{
    /// <summary>
    /// TypeDescriptor Services
    /// </summary>
    public static class TypeDescriptor
    {
        /// <summary>
        /// Returns the name of the class for the specified component using the default type descriptor.
        /// </summary>
        /// <param name="component">The System.Object for which you want the class name</param>
        /// <returns>A System.String containing the name of the class for the specified component.</returns>
        public static string GetClassName(object component)
        {
            return System.ComponentModel.TypeDescriptor.GetClassName(component);
        }

        /// <summary>
        /// Returns the name of the component for the specified component using the default type descriptor.
        /// </summary>
        /// <param name="component">The System.Object for which you want the component name</param>
        /// <returns>A System.String containing the name of the component for the specified component.</returns>
        public static string GetComponentName(object component)
        {
            return System.ComponentModel.TypeDescriptor.GetComponentName(component);
        }

        /// <summary>
        /// Returns the name of the component for the specified component using the default type descriptor.
        /// </summary>
        /// <param name="component">The System.Object for which you want the component name</param>
        /// <returns>A System.String containing the name of the component for the specified component.</returns>
        public static string GetFullComponentName(object component)
        {
            return System.ComponentModel.TypeDescriptor.GetFullComponentName(component);
        }

        /// <summary>
        /// Combines GetFullComponentName/GetClassName
        /// </summary>
        /// <param name="component">The System.Object for which you want the component+class name</param>
        /// <returns>A System.String containing the name of the component for the specified component.</returns>
        public static string GetFullComponentClassName(object component)
        {
            string name1 = System.ComponentModel.TypeDescriptor.GetFullComponentName(component);
            string name2 = System.ComponentModel.TypeDescriptor.GetClassName(component);
            return name1 + "." + name2;
        }

#if DEBUG
        /// <summary>
        /// Write GetFullComponentClassName
        /// </summary>
        /// <param name="console">console to write</param>
        /// <param name="message">message as any</param>
        /// <param name="component">The System.Object for which you want the component+class name</param>
        public static void WriteFullComponentClassName(DebugConsole console, string message, object component)
        {
            console.WriteLine(message, GetFullComponentClassName(component));
        }
#endif
    }
}
