using System;
using System.Runtime.InteropServices;

namespace NetOffice.Tools.Isolation
{
    /// <summary>
    ///  Represents an inner aggregator by a managed addin
    /// </summary>
    [ComImport, Guid("FBA7450D-B6E0-4E5C-908D-396BEFFC1D9B")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IManagedInnerAggregator
    {
        /// <summary>
        /// Creates a new instance of fullQualifiedTypeName and call outerObject to publish the instance
        /// to the outer caller.
        /// </summary>
        /// <param name="fullQualifiedTypeName">type to create</param>
        /// <param name="outerAggregator">caller as outer aggregator</param>
        /// <returns>true if inner aggregator accept the call, otherwise false</returns>
        bool CreateAggregatedInstance(string fullQualifiedTypeName, IOuterComAggregator outerAggregator);
    }
}
