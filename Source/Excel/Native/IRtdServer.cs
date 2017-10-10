using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi.Native
{
    /// <summary>
    /// NativeInterface IRtdServer
    /// Represents an interface for a real-time data server.
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [Guid("EC0E6191-DB51-11D3-8F3E-00C04F3651B8"), TypeLibType((short)4160)]
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsDual), ComVisible(true)]
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsNativeInterface)]
    public interface IRtdServer
	{
        #region Methods

        /// <summary>
        /// The ServerStart method is called immediately after a real-time data server is instantiated. Negative value or zero indicates failure to start the server; positive value indicates success.
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        /// <param name="callbackObject">IRTDUpdateEvent object. The callback object.</param>
        /// <returns>System.Int32</returns>
        [DispId(10), MethodImpl(4096)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        int ServerStart([MarshalAs(UnmanagedType.Interface)] [In] IRTDUpdateEvent callbackObject);

        /// <summary>
        /// Adds new topics from a real-time data server. The ConnectData method is called when a file is opened that contains real-time data functions or when a user types in a new formula which contains the RTD function.
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        /// <param name="topicID">a unique value, assigned by Microsoft Excel, which identifies the topic</param>
        /// <param name="strings">a single-dimensional array of strings identifying the topic</param>
        /// <param name="getNewValues"></param>
        /// <returns>true to determine if new values are to be acquired</returns>
        [DispId(11), MethodImpl(4096)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [return: MarshalAs(UnmanagedType.Struct)]
        object ConnectData([In] int topicID, [MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] [In] ref object strings, [In] [Out] ref bool getNewValues);

        /// <summary>
        /// This method is called by Microsoft Excel to get new data.
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        /// <param name="topicCount">the RTD server must change the value of the TopicCount to the number of elements in the array returned</param>
        /// <returns>System.Array</returns>
        /// <remarks>The data returned to Excel is an Object containing a two-dimensional array. The first dimension represents the list of topic IDs. The second dimension represents the values associated with the topic IDs</remarks>
        [DispId(12), MethodImpl(4096)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)]
        object RefreshData([In] [Out] ref int topicCount);

        /// <summary>
        /// Notifies a real-time data (RTD) server application that a topic is no longer in use.
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        /// <param name="topicID">a unique value assigned to the topic assigned by Microsoft Excel</param>
        [DispId(13), MethodImpl(4096)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void DisconnectData([In] int topicID);

        /// <summary>
        /// Determines if the real-time data server is still active. Zero or a negative number indicates failure; a positive number indicates that the server is active.
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        /// <returns>System.Int32</returns>
        /// <remarks>The Heartbeat method is called by Microsoft Excel if the HeartbeatInterval property has elapsed since the last time Excel was called with the UpdateNotify method</remarks>
        [DispId(14), MethodImpl(4096)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        int Heartbeat();

        /// <summary>
        /// Terminates the connection to the real-time data server.
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [DispId(15), MethodImpl(4096)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        void ServerTerminate();

        #endregion
    }
}