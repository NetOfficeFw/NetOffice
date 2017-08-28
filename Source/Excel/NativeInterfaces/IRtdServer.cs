using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    #pragma warning disable
    /// <summary>
    /// NativeInterface IRtdServer SupportByVersion Excel, 10,11,12,14,15,16
    /// Represents an interface for a real-time data server.
    /// </summary>
    [SupportByVersion("Excel", 10,11,12,14,15,16)]
	[ComImport, ComVisible(true), Guid("EC0E6191-DB51-11D3-8F3E-00C04F3651B8"), TypeLibType((short) 4160)]
	[EntityType(EntityType.IsNativeInterface)]
	public interface IRtdServer
	{
        #region Methods

        /// <summary>
        /// SupportByVersion Excel, 10,11,12,14,15,16
        /// The ServerStart method is called immediately after a real-time data server is instantiated. Negative value or zero indicates failure to start the server; positive value indicates success.
        /// </summary>
        /// <param name="CallbackObject">IRTDUpdateEvent object. The callback object.</param>
        /// <returns>System.Int32</returns>
        [SupportByVersion("Excel", 10,11,12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(10)]
		Int32 ServerStart([In, MarshalAs(UnmanagedType.Interface)] object CallbackObject);

        /// <summary>
        /// SupportByVersion Excel, 10,11,12,14,15,16
        /// Adds new topics from a real-time data server. The ConnectData method is called when a file is opened that contains real-time data functions or when a user types in a new formula which contains the RTD function.
        /// </summary>
        /// <param name="TopicID">A unique value, assigned by Microsoft Excel, which identifies the topic</param>
        /// <param name="Strings">A single-dimensional array of strings identifying the topic</param>
        /// <param name="GetNewValues">True to determine if new values are to be acquired</param>
        /// <returns>System.Object</returns>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		[return: MarshalAs(UnmanagedType.Struct)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(11)]
		object ConnectData([In]Int32 TopicID, [In] object Strings, [In]bool GetNewValues);

        /// <summary>
        /// SupportByVersion Excel, 10,11,12,14,15,16
        /// This method is called by Microsoft Excel to get new data.
        /// </summary>
        /// <param name="TopicCount">The RTD server must change the value of the TopicCount to the number of elements in the array returned</param>
        /// <returns>System.Array</returns>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(12)]
		object RefreshData([In]Int32 TopicCount);

        /// <summary>
        /// SupportByVersion Excel, 10,11,12,14,15,16
        /// Notifies a real-time data (RTD) server application that a topic is no longer in use
        /// </summary>
        /// <param name="TopicID">A unique value assigned to the topic assigned by Microsoft Excel</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(13)]
		void DisconnectData([In]Int32 TopicID);

        /// <summary>
        /// SupportByVersion Excel, 10,11,12,14,15,16
        /// Determines if the real-time data server is still active. Zero or a negative number indicates failure; a positive number indicates that the server is active
        /// </summary>
        /// <returns>System.Int32</returns>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(14)]
		Int32 Heartbeat();

        /// <summary>
        /// SupportByVersion Excel, 10,11,12,14,15,16
        /// Terminates the connection to the real-time data server.
        /// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(15)]
		void ServerTerminate();

		#endregion

		#region Properties

		#endregion
	}
    #pragma warning restore
}