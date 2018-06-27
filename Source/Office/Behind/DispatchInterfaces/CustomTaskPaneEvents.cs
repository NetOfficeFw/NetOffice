using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface CustomTaskPaneEvents 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class CustomTaskPaneEvents : COMObject, NetOffice.OfficeApi.CustomTaskPaneEvents
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
                    _contractType = typeof(NetOffice.OfficeApi.CustomTaskPaneEvents);
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
                    _type = typeof(CustomTaskPaneEvents);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public CustomTaskPaneEvents() : base()
		{

		}

		#endregion

        #region Properties

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="customTaskPaneInst">NetOffice.OfficeApi._CustomTaskPane customTaskPaneInst</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void VisibleStateChange(NetOffice.OfficeApi._CustomTaskPane customTaskPaneInst)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "VisibleStateChange", customTaskPaneInst);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="customTaskPaneInst">NetOffice.OfficeApi._CustomTaskPane customTaskPaneInst</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void DockPositionStateChange(NetOffice.OfficeApi._CustomTaskPane customTaskPaneInst)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DockPositionStateChange", customTaskPaneInst);
        }

        #endregion

        #pragma warning restore
    }
}
