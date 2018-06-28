using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface IRibbonUI 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remark> https://msdn.microsoft.com/de-de/library/aa433869(v=office.12).aspx </remark> 
    [SupportByVersion("Office", 12, 14, 15, 16), NativeCaller(typeof(Native.IRibbonUI))]
    [EntityType(EntityType.IsNativeInterfaceCaller)]
    [TypeId("000C03A7-0000-0000-C000-000000000046")]
    [NativeCallerWrapper(typeof(NetOffice.OfficeApi.Behind.IRibbonUI))]
    public interface IRibbonUI : ICOMObject
    {
        #region Properties

        /// <summary>
        /// Instance has native EarlyBind Interface instead of UnderlyingObject 
        /// </summary>
        bool HasUnderlyingObject { get; }

        #endregion

        #region Methods

        /// <summary>
        /// Invalidates the cached values for all of the controls of the Ribbon user interface.
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/de-de/library/aa433552(v=office.12).aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Invalidate();

        /// <summary>
        /// Invalidates the cached values for all of the controls of the Ribbon user interface.
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="controlID">Specified the identifier of the control that will be invalidated.</param>
        /// <remarks> https://msdn.microsoft.com/de-de/library/aa433553(v=office.12).aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void InvalidateControl(string controlID);

        /// <summary>
        /// Used to invalidate a built-in control.
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        /// <param name="controlID">Specified the identifier of the control that will be invalidated.</param>
        [SupportByVersion("Office", 14, 15, 16)]
        void InvalidateControlMso(string controlID);

        /// <summary>
        /// Activates the specified custom tab.
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        /// <param name="controlID">Specifies the identifier of the custom Ribbon tab to be activated</param>
        [SupportByVersion("Office", 14, 15, 16)]
        void ActivateTab(string controlID);

        /// <summary>
        /// Activates the specified built-in tab.
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        /// <param name="controlID">Specifies the identifier of the custom Ribbon tab to be activated.</param>
		[SupportByVersion("Office", 14, 15, 16)]
        void ActivateTabMso(string controlID);

        /// <summary>
        /// Activates the specified custom tab on the Microsoft Office Fluent Ribbon UI. Uses the fully qualified name of the tab which includes the identifier and the namespace of the tab.
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        /// <param name="controlID">Specifies the identifier of the custom Ribbon tab to be activated</param>
        /// <param name="_namespace">Specifies the namespace of the tab element</param>
        [SupportByVersion("Office", 14, 15, 16)]
        void ActivateTabQ(string controlID, string _namespace);

        #endregion
    }
}
