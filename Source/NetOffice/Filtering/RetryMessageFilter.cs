using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace NetOffice.Filtering
{
    /// <summary>
    /// An IMessageFilter Implementation
    /// Learn more about: http://netoffice.codeplex.com/wikipage?title=Settings.MessageFilter_EN
    /// </summary>
    [TypeConverter(typeof(ExpandableObjectConverter))]
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
        [Description("Get or set the message filter is enabled"), DefaultValue(false), Category("RetryMessageFilter")]
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
        [Description("Get or set retry options"), DefaultValue(typeof(RetryMessageFilterMode), "Immediately"), Category("RetryMessageFilter")]
        public RetryMessageFilterMode RetryMode { get; set; }

        /// <summary>
        /// Get or set log options
        /// </summary>
        [Description("Get or set log options"), DefaultValue(typeof(RetryMessageFilterLogMode), "None"), Category("RetryMessageFilter")]
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
                DebugConsole.Default.WriteLine("IMessageFilter.RetryRejectedCall.dwTickCount={0} , dwRejectType={1}", dwTickCount, dwRejectType);

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
                DebugConsole.Default.WriteLine("IMessageFilter.MessagePending.dwTickCount={0} , dwPendingType={1}", dwTickCount, dwPendingType);
            return 1; // PENDINGMSG_WAITNOPROCESS see: http://msdn.microsoft.com/en-us/library/aa908923.aspx for further info
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Returns a System.String that represents the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            return String.Format("Enabled:{0} Mode:{1}", Enabled, RetryMode);
        }

        #endregion
    }
}