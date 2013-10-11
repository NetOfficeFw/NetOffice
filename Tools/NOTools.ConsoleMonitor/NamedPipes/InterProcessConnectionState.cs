using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.ConsoleMonitor.NamedPipes
{
    #region Comments
    /// <summary>
    /// 
    /// </summary>
    #endregion
    public enum InterProcessConnectionState
    {
        #region Comments
        /// <summary>
        /// 
        /// </summary>
        #endregion
        NotSet = 0,
        #region Comments
        /// <summary>
        /// 
        /// </summary>
        #endregion
        Error = 1,
        #region Comments
        /// <summary>
        /// 
        /// </summary>
        #endregion
        Creating = 2,
        #region Comments
        /// <summary>
        /// 
        /// </summary>
        #endregion
        Created = 3,
        #region Comments
        /// <summary>
        /// 
        /// </summary>
        #endregion
        WaitingForClient = 4,
        #region Comments
        /// <summary>
        /// 
        /// </summary>
        #endregion
        ConnectedToClient = 5,
        #region Comments
        /// <summary>
        /// 
        /// </summary>
        #endregion
        ConnectingToServer = 6,
        #region Comments
        /// <summary>
        /// 
        /// </summary>
        #endregion
        ConnectedToServer = 7,
        #region Comments
        /// <summary>
        /// 
        /// </summary>
        #endregion
        Reading = 8,
        #region Comments
        /// <summary>
        /// 
        /// </summary>
        #endregion
        ReadData = 9,
        #region Comments
        /// <summary>
        /// 
        /// </summary>
        #endregion
        Writing = 10,
        #region Comments
        /// <summary>
        /// 
        /// </summary>
        #endregion
        WroteData = 11,
        #region Comments
        /// <summary>
        /// 
        /// </summary>
        #endregion
        Flushing = 12,
        #region Comments
        /// <summary>
        /// 
        /// </summary>
        #endregion
        FlushedData = 13,
        #region Comments
        /// <summary>
        /// 
        /// </summary>
        #endregion
        Disconnecting = 14,
        #region Comments
        /// <summary>
        /// 
        /// </summary>
        #endregion
        Disconnected = 15,
        #region Comments
        /// <summary>
        /// 
        /// </summary>
        #endregion
        Closing = 16,
        #region Comments
        /// <summary>
        /// 
        /// </summary>
        #endregion
        Closed = 17,
    }
}
