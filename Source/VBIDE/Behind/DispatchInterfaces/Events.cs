using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi.Behind
{
    /// <summary>
    /// DispatchInterface Events
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class Events : COMObject, NetOffice.VBIDEApi.Events
    {
        #pragma warning disable

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.VBIDEApi.Events);
                return _contractType;
            }
        }
        private static Type _contractType;


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

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Events);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Events() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        /// <param name="vBProject">NetOffice.VBIDEApi.VBProject vBProject</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.VBIDEApi.ReferencesEvents get_ReferencesEvents(NetOffice.VBIDEApi.VBProject vBProject)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.ReferencesEvents>(this, "ReferencesEvents", null, vBProject);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_ReferencesEvents
        /// </summary>
        /// <param name="vBProject">NetOffice.VBIDEApi.VBProject vBProject</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_ReferencesEvents")]
        public virtual NetOffice.VBIDEApi.ReferencesEvents ReferencesEvents(NetOffice.VBIDEApi.VBProject vBProject)
        {
            return get_ReferencesEvents(vBProject);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        /// <param name="commandBarControl">object commandBarControl</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.VBIDEApi.CommandBarEvents get_CommandBarEvents(object commandBarControl)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.CommandBarEvents>(this, "CommandBarEvents", typeof(NetOffice.VBIDEApi.CommandBarEvents), commandBarControl);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_CommandBarEvents
        /// </summary>
        /// <param name="commandBarControl">object commandBarControl</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_CommandBarEvents")]
        public virtual NetOffice.VBIDEApi.CommandBarEvents CommandBarEvents(object commandBarControl)
        {
            return get_CommandBarEvents(commandBarControl);
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}
