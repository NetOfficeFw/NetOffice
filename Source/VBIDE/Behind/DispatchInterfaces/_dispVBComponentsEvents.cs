using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi.Behind
{
    /// <summary>
    /// DispatchInterface _dispVBComponentsEvents
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class _dispVBComponentsEvents : COMObject, NetOffice.VBIDEApi._dispVBComponentsEvents
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
                    _contractType = typeof(NetOffice.VBIDEApi._dispVBComponentsEvents);
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
                    _type = typeof(_dispVBComponentsEvents);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _dispVBComponentsEvents() : base()
		{

		}

		#endregion

        #region Properties

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual void ItemAdded(NetOffice.VBIDEApi.VBComponent vBComponent)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ItemAdded", vBComponent);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual void ItemRemoved(NetOffice.VBIDEApi.VBComponent vBComponent)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ItemRemoved", vBComponent);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
        /// <param name="oldName">string oldName</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual void ItemRenamed(NetOffice.VBIDEApi.VBComponent vBComponent, string oldName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ItemRenamed", vBComponent, oldName);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual void ItemSelected(NetOffice.VBIDEApi.VBComponent vBComponent)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ItemSelected", vBComponent);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual void ItemActivated(NetOffice.VBIDEApi.VBComponent vBComponent)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ItemActivated", vBComponent);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual void ItemReloaded(NetOffice.VBIDEApi.VBComponent vBComponent)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ItemReloaded", vBComponent);
        }

        #endregion

        #pragma warning restore
    }
}
