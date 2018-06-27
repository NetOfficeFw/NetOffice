using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface ICustomXMLPartEvents 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class ICustomXMLPartEvents : COMObject, NetOffice.OfficeApi.ICustomXMLPartEvents
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
                    _contractType = typeof(NetOffice.OfficeApi.ICustomXMLPartEvents);
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
                    _type = typeof(ICustomXMLPartEvents);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ICustomXMLPartEvents() : base()
		{

		}

		#endregion

        #region Properties

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="newNode">NetOffice.OfficeApi.CustomXMLNode newNode</param>
        /// <param name="inUndoRedo">bool inUndoRedo</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void NodeAfterInsert(NetOffice.OfficeApi.CustomXMLNode newNode, bool inUndoRedo)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "NodeAfterInsert", newNode, inUndoRedo);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
        /// <param name="oldParentNode">NetOffice.OfficeApi.CustomXMLNode oldParentNode</param>
        /// <param name="oldNextSibling">NetOffice.OfficeApi.CustomXMLNode oldNextSibling</param>
        /// <param name="inUndoRedo">bool inUndoRedo</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void NodeAfterDelete(NetOffice.OfficeApi.CustomXMLNode oldNode, NetOffice.OfficeApi.CustomXMLNode oldParentNode, NetOffice.OfficeApi.CustomXMLNode oldNextSibling, bool inUndoRedo)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "NodeAfterDelete", oldNode, oldParentNode, oldNextSibling, inUndoRedo);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode oldNode</param>
        /// <param name="newNode">NetOffice.OfficeApi.CustomXMLNode newNode</param>
        /// <param name="inUndoRedo">bool inUndoRedo</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void NodeAfterReplace(NetOffice.OfficeApi.CustomXMLNode oldNode, NetOffice.OfficeApi.CustomXMLNode newNode, bool inUndoRedo)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "NodeAfterReplace", oldNode, newNode, inUndoRedo);
        }

        #endregion

        #pragma warning restore
    }
}
