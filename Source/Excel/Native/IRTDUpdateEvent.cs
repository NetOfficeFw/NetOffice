using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi.Native
{
    /// <summary>
    /// NativeInterface IRTDUpdateEvent
    /// Represents real-time data update events.
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
	[Guid("A43788C1-D91B-11D3-8F39-00C04F3651B8"), TypeLibType((short)4160)]
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsDual), ComVisible(true)]
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsNativeInterface)]
    public interface IRTDUpdateEvent
	{
        /// <summary>
        /// Returns or sets an Integer for the interval between updates for real-time data. Read/write.
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// Get/Set
        /// </summary>
        [DispId(11)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        int HeartbeatInterval
        {
            //SupportByVersion Excel, 9,10,11,12,14,15,16
            [DispId(11), MethodImpl(4096)]
            [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
            get;
            //SupportByVersion Excel, 9,10,11,12,14,15,16
            [DispId(11), MethodImpl(4096)]
            [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
            set;
        }

        /// <summary>
        /// The real-time data (RTD) server uses this method to notify Microsoft Excel that new data has been received.
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [DispId(10)]
        [MethodImpl(4096)]
        void UpdateNotify();

        /// <summary>
        /// Instructs the real-time data server (RTD) to disconnect from the specified IRTDUpdateEvent object.
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [DispId(12)]
        [MethodImpl(4096)]
        void Disconnect();
    }
}