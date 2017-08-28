using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    #pragma warning disable
    /// <summary>
    /// NativeInterface IRTDUpdateEvent SupportByVersion Excel, 10,11,12,14,15,16
    /// Represents real-time data update events.
    /// </summary>
    [SupportByVersion("Excel", 10,11,12,14,15,16)]
	[ComImport, ComVisible(true), Guid("A43788C1-D91B-11D3-8F39-00C04F3651B8"), TypeLibType((short) 4160)]
	[EntityType(EntityType.IsNativeInterface)]
	public interface IRTDUpdateEvent
	{
        #region Methods

        /// <summary>
        /// SupportByVersion Excel, 10,11,12,14,15,16
        /// The real-time data (RTD) server uses this method to notify Microsoft Excel that new data has been received.
        /// </summary>
        [SupportByVersion("Excel", 10,11,12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(10)]
		void UpdateNotify();

        /// <summary>
        /// SupportByVersion Excel, 10,11,12,14,15,16
        /// Instructs the real-time data server (RTD) to disconnect from the specified IRTDUpdateEvent object.
        /// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(12)]
		void Disconnect();

		#endregion

		#region Properties

		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		[DispId(11)]
		Int32 HeartbeatInterval{[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType=MethodCodeType.Runtime), DispId(11)] get;}

		#endregion
	}
    #pragma warning restore
}




