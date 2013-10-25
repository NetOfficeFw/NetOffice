using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Indicates in which method the error is occured
    /// </summary>
    public enum ErrorMethodKind
    {
        /// <summary>
        /// the error is occured in void IDTExtensibility2.OnStartupComplete(ref Array custom)
        /// </summary>
        OnStartupComplete = 0,

        /// <summary>
        /// the error is occured in void IDTExtensibility2.OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        /// </summary>
        OnDisconnection = 1,

        /// <summary>
        /// the error is occured in void IDTExtensibility2.OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        /// </summary>
        OnConnection = 2,

        /// <summary>
        ///  the error is occured in void IDTExtensibility2.OnAddInsUpdate(ref Array custom)
        /// </summary>
        OnAddInsUpdate = 3,

        /// <summary>
        /// the error is occured in void IDTExtensibility2.OnBeginShutdown(ref Array custom)
        /// </summary>
        OnBeginShutdown = 4,

        /// <summary>
        /// the error is occured in public virtual string GetCustomUI(string RibbonID)
        /// </summary>
        GetCustomUI = 5,

        /// <summary>
        /// the error is occured in public virtual void CTPFactoryAvailable(object CTPFactoryInst)
        /// </summary>
        CTPFactoryAvailable = 6,

        /// <summary>
        ///  the error is occured in protected virtual Factory CreateFactory()
        /// </summary>
        CreateFactory = 7

    }
}
