using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface IRibbonUI 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [IsNativeCallerWrapper(typeof(NetOffice.OfficeApi.IRibbonUI))]
    public class IRibbonUI : COMObject, NetOffice.OfficeApi.IRibbonUI
    {
        #region Type Information

        /// <summary>
        /// Instance Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type InstanceType
        {
            get
            {
                return LateBindingApiWrapperType;
            }
        }

        private static Type _type;

        /// <summary>
        /// Static Type Cache
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(IRibbonUI);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public IRibbonUI()
        {

        }

        #endregion

        #region Properties

        /// <summary>
        /// Instance has native EarlyBind Interface instead of UnderlyingObject 
        /// </summary>
        public bool HasUnderlyingObject
        {
            get
            {
                return null != NativeRibbon;
            }
        }

        /// <summary>
        /// Native EarlyBind Interface used instead of UnderlyingObject
        /// </summary>
        protected internal Native.IRibbonUI NativeRibbon { get; private set; }

        #endregion

        #region Overrides 

        /// <summary>
        /// Called from ctor at last as an inherited class service
        /// </summary>
        protected override void OnCreate()
        {
            base.OnCreate();
            NativeRibbon = UnderlyingObject as Native.IRibbonUI;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Invalidates the cached values for all of the controls of the Ribbon user interface.
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> https://msdn.microsoft.com/de-de/library/aa433552(v=office.12).aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Invalidate()
        {
            if (HasUnderlyingObject)
                NativeRibbon.Invalidate();
        }

        /// <summary>
        /// Invalidates the cached values for all of the controls of the Ribbon user interface.
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="controlID">Specified the identifier of the control that will be invalidated.</param>
        /// <remarks> https://msdn.microsoft.com/de-de/library/aa433553(v=office.12).aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void InvalidateControl(string controlID)
        {
            if (HasUnderlyingObject)
                NativeRibbon.InvalidateControl(controlID);
        }

        /// <summary>
        /// Used to invalidate a built-in control.
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        /// <param name="controlID">Specified the identifier of the control that will be invalidated.</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual void InvalidateControlMso(string controlID)
        {
            if (HasUnderlyingObject)
                NativeRibbon.InvalidateControlMso(controlID);
        }

        /// <summary>
        /// Activates the specified custom tab.
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        /// <param name="controlID">Specifies the identifier of the custom Ribbon tab to be activated</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual void ActivateTab(string controlID)
        {
            if (HasUnderlyingObject)
                NativeRibbon.ActivateTab(controlID);
        }

        /// <summary>
        /// Activates the specified built-in tab.
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        /// <param name="controlID">Specifies the identifier of the custom Ribbon tab to be activated.</param>
		[SupportByVersion("Office", 14, 15, 16)]
        public virtual void ActivateTabMso(string controlID)
        {
            if (HasUnderlyingObject)
                NativeRibbon.ActivateTabMso(controlID);
        }

        /// <summary>
        /// Activates the specified custom tab on the Microsoft Office Fluent Ribbon UI. Uses the fully qualified name of the tab which includes the identifier and the namespace of the tab.
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        /// <param name="controlID">Specifies the identifier of the custom Ribbon tab to be activated</param>
        /// <param name="_namespace">Specifies the namespace of the tab element</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual void ActivateTabQ(string controlID, string _namespace)
        {
            if (HasUnderlyingObject)
                NativeRibbon.ActivateTabQ(controlID, _namespace);
        }

        #endregion
    }
}
