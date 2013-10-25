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
        /// The add-in was loaded after Application started.
        /// </summary>
        ext_cm_AfterStartup,

        /// <summary>
        /// The add-in was loaded when Application started.
        /// </summary>
        ext_cm_Startup,

        /// <summary>
        /// The add-in was loaded by an external client.
        /// </summary>
        ext_cm_External,

        /// <summary>
        /// The add-in was loaded from the command line.
        /// </summary>
        ext_cm_CommandLine,

        /// <summary>
        /// The add-in was loaded with a solution.
        /// </summary>
        ext_cm_Solution,

        /// <summary>
        /// The add-in was loaded for user interface setup.
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
        /// The add-in was unloaded when Application was shut down.
        /// </summary>
        ext_dm_HostShutdown,

        /// <summary>
        /// The add-in was unloaded while Application was running.
        /// </summary>
        ext_dm_UserClosed,

        /// <summary>
        /// The add-in was unloaded after the user interface was set up.
        /// </summary>
        ext_dm_UISetupComplete,

        /// <summary>
        /// The add-in was unloaded when the solution was closed.
        /// </summary>
        ext_dm_SolutionClosed
    }

    /// <summary>
    /// 
    /// </summary>
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), TypeLibType(4160), ComImport]
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
