using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace NetOffice.Tools
{
    /// <summary>
    /// 
    /// </summary>
    [Guid("289E9AF1-4973-11D1-AE81-00A0C90F26F4")]
    public enum ext_ConnectMode
    {
        /// <summary>
        /// 
        /// </summary>
        ext_cm_AfterStartup,

        /// <summary>
        /// 
        /// </summary>
        ext_cm_Startup,

        /// <summary>
        /// 
        /// </summary>
        ext_cm_External,

        /// <summary>
        /// 
        /// </summary>
        ext_cm_CommandLine,

        /// <summary>
        /// 
        /// </summary>
        ext_cm_Solution,

        /// <summary>
        /// 
        /// </summary>
        ext_cm_UISetup
    }

    /// <summary>
    /// 
    /// </summary>
    [Guid("289E9AF2-4973-11D1-AE81-00A0C90F26F4")]
    public enum ext_DisconnectMode
    {
        /// <summary>
        /// 
        /// </summary>
        ext_dm_HostShutdown,

        /// <summary>
        /// 
        /// </summary>
        ext_dm_UserClosed,

        /// <summary>
        /// 
        /// </summary>
        ext_dm_UISetupComplete,

        /// <summary>
        /// 
        /// </summary>
        ext_dm_SolutionClosed
    }

    /// <summary>
    /// 
    /// </summary>
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), TypeLibType(4160)]
    [ComImport]
    public interface IDTExtensibility2
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Application"></param>
        /// <param name="ConnectMode"></param>
        /// <param name="AddInInst"></param>
        /// <param name="custom"></param>
        [DispId(1)]
        [MethodImpl(4096)]
        void OnConnection([MarshalAs(26)] [In] object Application, [In] ext_ConnectMode ConnectMode, [MarshalAs(26)] [In] object AddInInst, [MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] [In] ref Array custom);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="RemoveMode"></param>
        /// <param name="custom"></param>
        [DispId(2)]
        [MethodImpl(4096)]
        void OnDisconnection([In] ext_DisconnectMode RemoveMode, [MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] [In] ref Array custom);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="custom"></param>
        [DispId(3)]
        [MethodImpl(4096)]
        void OnAddInsUpdate([MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] [In] ref Array custom);
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="custom"></param>
        [DispId(4)]
        [MethodImpl(4096)]
        void OnStartupComplete([MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] [In] ref Array custom);
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="custom"></param>
        [DispId(5)]
        [MethodImpl(4096)]
        void OnBeginShutdown([MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] [In] ref Array custom);
    }
}
