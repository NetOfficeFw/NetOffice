using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Globalization;
using System.Reflection;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Threading;

namespace NetOffice
{
    /// <summary>
    /// 
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct INTERFACEINFO
    {
        /// <summary>
        /// 
        /// </summary>
        [MarshalAs(UnmanagedType.IUnknown)]
        public object punk;
        
        /// <summary>
        /// 
        /// </summary>
        public Guid iid;
        
        /// <summary>
        /// 
        /// </summary>
        public ushort wMethod;
    }

    /// <summary>
    /// 
    /// </summary>
    [ComImport, ComConversionLoss, InterfaceType((short)1), Guid("00000016-0000-0000-C000-000000000046")]
    public interface IMessageFilter
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dwCallType"></param>
        /// <param name="htaskCaller"></param>
        /// <param name="dwTickCount"></param>
        /// <param name="lpInterfaceInfo"></param>
        /// <returns></returns>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        int HandleInComingCall([In] uint dwCallType, [In] IntPtr htaskCaller, [In] uint dwTickCount, [In, MarshalAs(UnmanagedType.LPArray)] INTERFACEINFO[] lpInterfaceInfo);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="htaskCallee"></param>
        /// <param name="dwTickCount"></param>
        /// <param name="dwRejectType"></param>
        /// <returns></returns>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        int RetryRejectedCall([In] IntPtr htaskCallee, [In] uint dwTickCount, [In] uint dwRejectType);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="htaskCallee"></param>
        /// <param name="dwTickCount"></param>
        /// <param name="dwPendingType"></param>
        /// <returns></returns>
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        int MessagePending([In] IntPtr htaskCallee, [In] uint dwTickCount, [In] uint dwPendingType);
    }

    /// <summary>
    /// Specify the filter for an RetryMessageFilter instance
    /// </summary>
    public enum RetryMessageFilterMode
    {
        /// <summary>
        /// Try rejected call again immediately
        /// </summary>
        Immediately = 0,

        /// <summary>
        /// Try rejected call again after few milliseconds
        /// </summary>
        Delayed = 1,

        /// <summary>
        /// Dont try rejected call again
        /// </summary>
        None = 2
    }

    /// <summary>
    /// Specify log behaviour for an RetryMessageFilter instance
    /// </summary>
    public enum RetryMessageFilterLogMode
    {
        /// <summary>
        /// Disable Log
        /// </summary>
        None = 0,
        
        /// <summary>
        /// Call DebugConsole.WriteLine in IMessageFilter.RetryRejectedCall
        /// </summary>
        RetryRejectedCall = 1,
        
        /// <summary>
        /// Call DebugConsole.WriteLine in IMessageFilter.MessagePending
        /// </summary>
        MessagePending = 2,
        
        /// <summary>
        /// Call DebugConsole.WriteLine in IMessageFilter.RetryRejectedCall and IMessageFilter.MessagePending
        /// </summary>
        Both = 3
    }

    /// <summary>
    /// An IMessageFilter Implementation
    /// Learn more about: http://netoffice.codeplex.com/wikipage?title=Settings.MessageFilter_EN
    /// </summary>
    public class RetryMessageFilter : IMessageFilter
    {
        #region Fields / Imports

        [DllImport("ole32.dll")]
        static extern int CoRegisterMessageFilter(IMessageFilter lpMessageFilter, out IMessageFilter lplpMessageFilter);

        private IMessageFilter _messageFilter;

        #endregion

        #region Properties

        /// <summary>
        /// Get or set the message filter is enabled
        /// </summary>
        public bool Enabled
        {
            get
            {
                return (_messageFilter != null);
            }
            set
            {
                if (value)
                    RegisterFilter();
                else
                    UnregisterFilter();
            }
        }

        /// <summary>
        /// Get or set retry options
        /// </summary>
        public RetryMessageFilterMode RetryMode { get; set; }

        /// <summary>
        /// Get or set log options
        /// </summary>
        public RetryMessageFilterLogMode LogMode { get; set; }

        #endregion

        #region Methods

        private void RegisterFilter()
        {
            CoRegisterMessageFilter(this, out _messageFilter);
        }

        private void UnregisterFilter()
        {
            _messageFilter = null;
            CoRegisterMessageFilter(null, out _messageFilter);
        }

        #endregion

        #region IMessageFilter Member

        int IMessageFilter.HandleInComingCall(uint dwCallType, IntPtr htaskCaller, uint dwTickCount, INTERFACEINFO[] lpInterfaceInfo)
        {
            return 1;  // SERVERCALL_REJECTED - We're the client, so we won't get HandleInComingCall calls.
        }

        int IMessageFilter.RetryRejectedCall(IntPtr htaskCallee, uint dwTickCount, uint dwRejectType)
        {
            if (LogMode == RetryMessageFilterLogMode.RetryRejectedCall || LogMode == RetryMessageFilterLogMode.Both)
                DebugConsole.WriteLine("IMessageFilter.RetryRejectedCall.dwTickCount={0} , dwRejectType={1}", dwTickCount, dwRejectType);

            switch (RetryMode)
            {
                case RetryMessageFilterMode.Immediately:
                    return 1;
                case RetryMessageFilterMode.Delayed:
                    return 101;
                case RetryMessageFilterMode.None:
                    return -1;
                default:
                    throw new IndexOutOfRangeException("RetryMessageFilterMode");
            }
        }

        int IMessageFilter.MessagePending(IntPtr htaskCallee, uint dwTickCount, uint dwPendingType)
        {
            if (LogMode == RetryMessageFilterLogMode.MessagePending || LogMode == RetryMessageFilterLogMode.Both)
                DebugConsole.WriteLine("IMessageFilter.MessagePending.dwTickCount={0} , dwPendingType={1}", dwTickCount, dwPendingType);
            return 1; // PENDINGMSG_WAITNOPROCESS see: http://msdn.microsoft.com/en-us/library/aa908923.aspx for further info
        }

        #endregion

    }
}