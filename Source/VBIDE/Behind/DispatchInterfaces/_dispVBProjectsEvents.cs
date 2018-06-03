using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi.Behind
{
    /// <summary>
    /// DispatchInterface _dispVBProjectsEvents
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class _dispVBProjectsEvents : COMObject, NetOffice.VBIDEApi.X_dispVBProjectsEvents
    {
        #pragma warning disable

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

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(_dispVBProjectsEvents);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _dispVBProjectsEvents() : base()
		{

		}

		#endregion

        #region Properties

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBProject">NetOffice.VBIDEApi.VBProject vBProject</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public void ItemAdded(NetOffice.VBIDEApi.VBProject vBProject)
        {
            Factory.ExecuteMethod(this, "ItemAdded", vBProject);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBProject">NetOffice.VBIDEApi.VBProject vBProject</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public void ItemRemoved(NetOffice.VBIDEApi.VBProject vBProject)
        {
            Factory.ExecuteMethod(this, "ItemRemoved", vBProject);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBProject">NetOffice.VBIDEApi.VBProject vBProject</param>
        /// <param name="oldName">string oldName</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public void ItemRenamed(NetOffice.VBIDEApi.VBProject vBProject, string oldName)
        {
            Factory.ExecuteMethod(this, "ItemRenamed", vBProject, oldName);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="vBProject">NetOffice.VBIDEApi.VBProject vBProject</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public void ItemActivated(NetOffice.VBIDEApi.VBProject vBProject)
        {
            Factory.ExecuteMethod(this, "ItemActivated", vBProject);
        }

        #endregion

        #pragma warning restore
    }
}
