using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace NetOffice.Tools
{
    /// <summary>
    /// Used in IDTExtensibility2 interface
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
    /// Used in IDTExtensibility2 interface
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
    /// The well known Extensibility
    /// </summary>
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), TypeLibType(4160), ComImport]
    public interface IDTExtensibility2
    {
        /// <summary>
        /// Occurs whenever an add-in is loaded into MS-Office
        /// </summary>
        /// <param name="Application">A reference to an instance of the office application</param>
        /// <param name="ConnectMode">An ext_ConnectMode enumeration value that indicates the way the add-in was loaded into MS-Office</param>
        /// <param name="AddInInst">An AddIn reference to the add-in's own instance. This is stored for later use, such as determining the parent collection for the add-in</param>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use in the add-in</param>
        [DispId(1)]
        [MethodImpl(4096)]
        void OnConnection([MarshalAs(26)] [In] object Application, [In] ext_ConnectMode ConnectMode, [MarshalAs(26)] [In] object AddInInst, [MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] [In] ref Array custom);

        /// <summary>
        /// Occurs whenever an add-in is unloaded from MS Office
        /// </summary>
        /// <param name="RemoveMode">An ext_DisconnectMode enumeration value that informs an add-in why it was unloaded.</param>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use after the add-in unloads</param>
        [DispId(2)]
        [MethodImpl(4096)]
        void OnDisconnection([In] ext_DisconnectMode RemoveMode, [MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] [In] ref Array custom);

        /// <summary>
        /// Occurs whenever an add-in is loaded or unloaded from MS Office
        /// </summary>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use in the add-in</param>
        [DispId(3)]
        [MethodImpl(4096)]
        void OnAddInsUpdate([MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] [In] ref Array custom);

        /// <summary>
        ///  Occurs whenever an add-in, which is set to load when MS Office starts, loads.
        /// </summary>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use when the add-in loads</param>
        [DispId(4)]
        [MethodImpl(4096)]
        void OnStartupComplete([MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] [In] ref Array custom);

        /// <summary>
        /// Occurs whenever MS Office shuts down while an add-in is running
        /// </summary>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use in the add-in</param>
        [DispId(5)]
        [MethodImpl(4096)]
        void OnBeginShutdown([MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] [In] ref Array custom);
    }
}
