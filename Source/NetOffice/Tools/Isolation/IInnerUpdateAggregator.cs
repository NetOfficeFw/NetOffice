using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace NetOffice.Tools.Isolation
{
    /// <summary>
    /// To implement by a managed addin to recieve an update aggregator
    /// </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport, Guid("EF261BCD-3078-459E-9448-13845BEED136")]
    public interface IInnerUpdateAggregator
    {
        /// <summary>
        /// Set an unmanaged aggregator to a managed addin instance
        /// </summary>
        /// <param name="aggregator">outer aggregator</param>
        void SetOuterAggregator([In, Out, MarshalAs(UnmanagedType.IUnknown)] IOuterUpdateAggregator aggregator);
    }
}
