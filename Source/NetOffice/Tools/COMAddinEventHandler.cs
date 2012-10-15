using System;

namespace NetOffice.Tools
{
    /// <summary>
    /// 
    /// </summary>
    /// <param name="custom"></param>
    public delegate void OnStartupCompleteEventHandler(ref Array custom);
    
    /// <summary>
    /// 
    /// </summary>
    /// <param name="RemoveMode"></param>
    /// <param name="custom"></param>
    public delegate void OnDisconnectionEventHandler(ext_DisconnectMode RemoveMode, ref Array custom);
    
    /// <summary>
    /// 
    /// </summary>
    /// <param name="Application"></param>
    /// <param name="ConnectMode"></param>
    /// <param name="AddInInst"></param>
    /// <param name="custom"></param>
    public delegate void OnConnectionEventHandler(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom);
    
    /// <summary>
    /// 
    /// </summary>
    /// <param name="custom"></param>
    public delegate void OnAddInsUpdateEventHandler(ref Array custom);
    
    /// <summary>
    /// 
    /// </summary>
    /// <param name="custom"></param>
    public delegate void OnBeginShutdownEventHandler(ref Array custom);
}
