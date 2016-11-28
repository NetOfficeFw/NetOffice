using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// Represents various dispose possibilities to free child instances
    /// </summary>
    public interface ICOMObjectTableDisposable
    {
        /// <summary>
        /// Dispose all child instance
        /// </summary>
        void DisposeChildInstances();


        /// <summary>
        /// Dispose all child instance
        /// </summary>
        /// <param name="disposeEventBinding">dispose open event subscriptions</param>
        void DisposeChildInstances(bool disposeEventBinding);
    }
}
