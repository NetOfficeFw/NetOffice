using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.CoreServices
{
    /// <summary>
    /// ProxyCountChanged delegate
    /// </summary>
    /// <param name="sender">sender instance</param>
    /// <param name="proxyCount">current count of com proxies</param>
    public delegate void CountChangedHandler(ICoreManagement sender, int proxyCount);

    /// <summary>
    /// Proxy added delegate
    /// </summary>
    /// <param name="sender">sender instance</param>
    /// <param name="ownerPath">comObject relation path</param>
    /// <param name="comObject">added object</param>
    public delegate void AddedHandler(ICoreManagement sender, IEnumerable<ICOMObject> ownerPath, ICOMObject comObject);

    /// <summary>
    /// Proxy remove delegate
    /// </summary>
    /// <param name="sender">sender instance</param>
    /// <param name="ownerPath">former comObject relation path</param>
    /// <param name="comObject">removed object</param>
    public delegate void RemovedHandler(ICoreManagement sender, IEnumerable<ICOMObject> ownerPath, ICOMObject comObject);

    /// <summary>
    /// Proxy clear delegate
    /// </summary>
    /// <param name="sender">sender instance</param>
    public delegate void ClearHandler(ICoreManagement sender);
}
