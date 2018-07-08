/*
    This part of the code is inspired by the legendary COM Shim Wizard.
    Credits goes to Garry Trinder and Misha Shneerson.
 */

using System;
using System.IO;
using System.Runtime.InteropServices;

namespace NetOffice.Tools.Isolation
{
    /// <summary>
    ///  Represents an inner aggregator by a managed update handler
    /// </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport, Guid("F249CE55-8BC5-44AB-81A0-1040998B6CD5")]
    public interface IManagedInnerUpdateAggregator
    {
        /// <summary>
        /// Creates a new instance of fullQualifiedTypeName and call outerObject to publish the instance
        /// to the outer caller.
        /// </summary>
        /// <param name="assemblyName">name or strong name where the target type is located</param>
        /// <param name="fullQualifiedTypeName">type to create</param>
        /// <param name="outerAggregator">caller as outer aggregator</param>
        /// <param name="outerUpdateAggregator">outer update aggregator</param>
        /// <exception cref="ArgumentNullException">argument is null or empty</exception>
        /// <exception cref="MissingMethodException">no matching public constructor was found</exception>/param>
        /// <exception cref="TypeLoadException">typename was not found in assemblyName</exception>/param>
        /// <exception cref="FileNotFoundException">assemblyName was not found</exception>/param>
        /// <exception cref="MethodAccessException">the caller does not have permission to call this constructor</exception>/param>
        /// <exception cref="AppDomainUnloadedException">the operation is attempted on an unloaded application domain.</exception>/param>
        /// <exception cref="BadImageFormatException">assemblyName is not a valid assembly</exception>
        /// <exception cref="FileLoadException">an assembly or module was loaded twice with two different evidences</exception>
        /// <exception cref="InvalidOperationException">unexpected error</exception>
        void CreateAggregatedInstance(string assemblyName, string fullQualifiedTypeName, IOuterUpdateAggregator outerAggregator, IShimUpdateHost outerUpdateAggregator);
    }
}
