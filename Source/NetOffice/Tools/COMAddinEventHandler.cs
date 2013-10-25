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
    /// <param name="RemoveMode">kind of remove</param>
    /// <param name="custom">custom arguments</param>
    public delegate void OnDisconnectionEventHandler(ext_DisconnectMode RemoveMode, ref Array custom);

    /// <summary>
    /// OnConnection Event Handler
    /// </summary>
    /// <param name="Application">application host instance</param>
    /// <param name="ConnectMode">kind of connect</param>
    /// <param name="AddInInst">addin instance</param>
    /// <param name="custom">custom arguments</param>
    public delegate void OnConnectionEventHandler(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom);

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
