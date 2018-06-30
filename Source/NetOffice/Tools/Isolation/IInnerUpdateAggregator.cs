using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Tools.Isolation
{
    /// <summary>
    /// To implement by a managed addin to recieve an update aggregator
    /// </summary>
    public interface IInnerUpdateAggregator
    {
        /// <summary>
        /// Set an unmanaged aggregator to a managed addin instance
        /// </summary>
        /// <param name="aggregator">outer aggregator</param>
        void SetOuterAggregator(IOuterUpdateAggregator aggregator);
    }
}
