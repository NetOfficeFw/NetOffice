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
    /// Default implementation of <see cref="NetOffice.Tools.Isolation.IManagedInnerAggregator"/>
    /// </summary>
    [ClassInterface(ClassInterfaceType.None)]
    public class ManagedInnerAggregator : IManagedInnerAggregator
    {
        /// <summary>
        /// Creates a new instance of fullQualifiedTypeName and call outerObject to publish the instance
        /// to the outer caller.
        /// </summary>
        /// <param name="assemblyName">name or strong name where the target type is located</param>
        /// <param name="fullQualifiedTypeName">type to create</param>
        /// <param name="outerAggregator">caller as outer aggregator</param>
        /// <exception cref="ArgumentNullException">argument is null or empty</exception>
        /// <exception cref="MissingMethodException">no matching public constructor was found</exception>/param>
        /// <exception cref="TypeLoadException">typename was not found in assemblyName</exception>/param>
        /// <exception cref="FileNotFoundException">assemblyName was not found</exception>/param>
        /// <exception cref="MethodAccessException">the caller does not have permission to call this constructor</exception>/param>
        /// <exception cref="AppDomainUnloadedException">the operation is attempted on an unloaded application domain.</exception>/param>
        /// <exception cref="BadImageFormatException">assemblyName is not a valid assembly</exception>
        /// <exception cref="FileLoadException">an assembly or module was loaded twice with two different evidences</exception>
        /// <exception cref="InvalidOperationException">unexpected error</exception>
        public void CreateAggregatedInstance(string assemblyName, string fullQualifiedTypeName, IOuterComAggregator outerAggregator)
        {
            if (String.IsNullOrWhiteSpace(assemblyName))
                throw new ArgumentNullException("assemblyName");
            if (String.IsNullOrWhiteSpace(fullQualifiedTypeName))
                throw new ArgumentNullException("fullQualifiedTypeName");
            if (null == outerAggregator)
                throw new ArgumentNullException("outerAggregator");

            IntPtr pOuter = IntPtr.Zero;
            IntPtr pInner = IntPtr.Zero;

            try
            {
                pOuter = Marshal.GetIUnknownForObject(outerAggregator);
                if (IntPtr.Zero != pOuter)
                {
                    object innerObject = AppDomain.CurrentDomain.CreateInstanceAndUnwrap(assemblyName, fullQualifiedTypeName);
                    pInner = Marshal.CreateAggregatedObject(pOuter, innerObject);
                    if (IntPtr.Zero != pInner)
                        outerAggregator.SetInnerAddin(pInner);
                }

                if (IntPtr.Zero == pOuter || IntPtr.Zero == pInner)
                    throw new InvalidOperationException("Unexpected result while getting IUnkown Pointer.");
            }
            catch(Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                throw;
            }
            finally
            {
                if (IntPtr.Zero  != pOuter)
                    Marshal.Release(pOuter);
                if (IntPtr.Zero  != pInner)
                    Marshal.Release(pInner);
                Marshal.ReleaseComObject(outerAggregator);
            }
        }
    }
}
