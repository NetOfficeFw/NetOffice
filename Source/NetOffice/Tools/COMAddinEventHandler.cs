using System;

namespace NetOffice.Tools
{
    /// <summary>
    /// OnStartupComplete Event Handler
    /// </summary>
    /// <param name="custom">custom arguments</param>
    public delegate void OnStartupCompleteEventHandler(ref Array custom);

    /// <summary>
    /// OnDisconnection Event Handler
    /// </summary>
    /// <param name="removeMode">kind of remove</param>
    /// <param name="custom">custom arguments</param>
    public delegate void OnDisconnectionEventHandler(ext_DisconnectMode removeMode, ref Array custom);

    /// <summary>
    /// OnConnection Event Handler
    /// </summary>
    /// <param name="application">application host instance</param>
    /// <param name="connectMode">kind of connect</param>
    /// <param name="addInInst">addin instance</param>
    /// <param name="custom">custom arguments</param>
    public delegate void OnConnectionEventHandler(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom);

    /// <summary>
    /// OnAddInsUpdate Event Handler
    /// </summary>
    /// <param name="custom">custom arguments</param>
    public delegate void OnAddInsUpdateEventHandler(ref Array custom);

    /// <summary>
    /// OnBeginShutdown Event Handler
    /// </summary>
    /// <param name="custom">custom arguments</param>
    public delegate void OnBeginShutdownEventHandler(ref Array custom);
}
