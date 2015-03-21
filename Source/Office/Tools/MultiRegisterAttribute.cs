using System;

namespace NetOffice.OfficeApi.Tools
{
    /// <summary>
    /// Set a target office product for registering
    /// </summary>
    public enum RegisterIn
    { 
        /// <summary>
        /// MS Excel in any version
        /// </summary>
        Excel,

        /// <summary>
        /// MS Word in any version
        /// </summary>
        Word,

        /// <summary>
        /// MS Outlook in any version
        /// </summary>
        Outlook,
        
        /// <summary>
        /// MS PowerPoint in any version
        /// </summary>
        PowerPoint,

        /// <summary>
        /// MS Access in any version
        /// </summary>
        Access,

        /// <summary>
        /// MS Project in any version
        /// </summary>
        MSProject,

        /// <summary>
        /// MS Visio in any version
        /// </summary>
        Visio
    }

    /// <summary>
    /// This attribute can be used from NetOffice.OfficeApi.Tools.COMAddin to specify multipe office products
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Class)]
    public class MultiRegisterAttribute : System.Attribute
    {
        /// <summary>
        /// The office products for addin registration
        /// </summary>
        public readonly RegisterIn[] Products;

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        /// <param name="products">The office products for addin registration</param>
        public MultiRegisterAttribute(params RegisterIn[] products)
        {
            Products = products;
        }

		/// <summary>
        /// Looks for the MultiRegisterAttribute. Throws an exception if not found
        /// </summary>
        /// <param name="type">the type you want looking for the attribute</param>
        /// <returns>MultiRegisterAttribute</returns>
		internal static MultiRegisterAttribute GetAttribute(Type type)
		{
		    object[] array = type.GetCustomAttributes(typeof(MultiRegisterAttribute), false);
            if (array.Length == 0)
                throw new ArgumentNullException("MultiRegisterAttribute is missing");
            return array[0] as MultiRegisterAttribute;
		}

        /// <summary>
        /// Get registry value string
        /// </summary>
        /// <param name="register">target office application</param>
        /// <returns>registry sub string from office application</returns>
        internal static string RegistryEntry(RegisterIn register)
        {
            if (register == RegisterIn.MSProject)
                return "MS Project";
            else
                return register.ToString();
        }
    }
}
