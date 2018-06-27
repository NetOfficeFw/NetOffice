/*
    This part of the code is inspired by the legendary COM Shim Wizard.
    Credits goes to Garry Trinder and Misha Shneerson. (MSFT)
 */

using System;
using System.Runtime.InteropServices;

namespace NetOffice.Tools.Isolation
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.Tools.Isolation.IManagedInnerAggregator"/>
    /// </summary>
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class ManagedInnerAggregator : IManagedInnerAggregator
    {
        /// <summary>
        /// Creates a new instance of fullQualifiedTypeName and call outerObject to publish the instance
        /// to the outer aggregator.
        /// </summary>
        /// <param name="fullQualifiedTypeName">type to create</param>
        /// <param name="outerAggregator">caller as outer aggregator</param>
        /// <returns>true if inner aggregator accept the call, otherwise false</returns>
        public bool CreateAggregatedInstance(string fullQualifiedTypeName, IOuterComAggregator outerAggregator)
        {
            bool result = false;
            IntPtr pOuter = IntPtr.Zero;
            IntPtr pInner = IntPtr.Zero;

            try
            {
                if (!String.IsNullOrWhiteSpace(fullQualifiedTypeName) && null != outerAggregator)
                {
                    pOuter = Marshal.GetIUnknownForObject(outerAggregator);
                    var assembly = Type.GetType(fullQualifiedTypeName).Assembly;
                    object innerObject = AppDomain.CurrentDomain.CreateInstanceAndUnwrap(assembly.FullName, fullQualifiedTypeName);
                    pInner = Marshal.CreateAggregatedObject(pOuter, innerObject);
                    outerAggregator.SetInnerAddin(pInner);
                    result = true;
                }
            }
            catch(Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
            }
            finally
            {
                if (pOuter != IntPtr.Zero)
                    Marshal.Release(pOuter);
                if (pInner != IntPtr.Zero)
                    Marshal.Release(pInner);
                Marshal.ReleaseComObject(outerAggregator);
            }

            return result;
        }
    }
}
