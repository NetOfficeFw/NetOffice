using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// specifiy possible registry locations
    /// </summary>
    public enum RegistrySaveLocation
    {
        /// <summary>
        /// CurrentUser Key
        /// </summary>
        CurrentUser = 0,

        /// <summary>
        /// LocalMachineKey (permissions required)
        /// </summary>
        LocalMachine = 1
    }

    /// <summary>
    /// Specify the addin registry keys for office was created in the Machine key or current user
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Class)]
    public class RegistryLocationAttribute : System.Attribute
    {
        /// <summary>
        /// Registry Location
        /// </summary>
        public readonly RegistrySaveLocation Value;

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        /// <param name="value">Registry location</param>
        public RegistryLocationAttribute(RegistrySaveLocation value)
        {
            Value = value;
        }
    }
}
