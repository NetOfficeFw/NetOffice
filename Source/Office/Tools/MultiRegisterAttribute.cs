﻿using System;

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
        Excel = 0,

        /// <summary>
        /// MS Word in any version
        /// </summary>
        Word = 1,

        /// <summary>
        /// MS Outlook in any version
        /// </summary>
        Outlook = 2,
        
        /// <summary>
        /// MS PowerPoint in any version
        /// </summary>
        PowerPoint = 3,

        /// <summary>
        /// MS Access in any version
        /// </summary>
        Access = 4,

        /// <summary>
        /// MS Visio in any version
        /// </summary>
        Visio = 6,

        /// <summary>
        /// MS Publisher
        /// </summary>
        Publisher = 7,
    }

    /// <summary>
    /// This attribute must be used for NetOffice.OfficeApi.Tools.COMAddin to specify multipe office products you want support
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Class, AllowMultiple=false)]
    public class MultiRegisterAttribute : System.Attribute
    {
        /// <summary>
        /// The office products for addin (un-)registration
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
        /// <param name="throwException">throw exception if not found</param>
        /// <returns>MultiRegisterAttribute instance</returns>
        internal static MultiRegisterAttribute GetAttribute(Type type, bool throwException = true)
		{
		    object[] array = type.GetCustomAttributes(typeof(MultiRegisterAttribute), false);
            if (array.Length == 0 && throwException)
                throw new ArgumentException("MultiRegisterAttribute is missing");
            if (array.Length > 0)
                return array[0] as MultiRegisterAttribute;
            else
                return null;
		}

        /// <summary>
        /// Get registry value string
        /// </summary>
        /// <param name="register">target office application</param>
        /// <returns>registry sub string from office application</returns>
        internal static string RegistryEntry(RegisterIn register)
        {
            return register.ToString();
        }
    }
}
