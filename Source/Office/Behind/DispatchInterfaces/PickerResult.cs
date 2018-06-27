using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface PickerResult 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861756.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class PickerResult : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.PickerResult
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
                    _contractType = typeof(NetOffice.OfficeApi.PickerResult);
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
                    _type = typeof(PickerResult);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public PickerResult() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863784.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual string Id
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Id");
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861059.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual string DisplayName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DisplayName");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayName", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865231.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual string Type
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Type");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Type", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863831.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual string SIPId
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "SIPId");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SIPId", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865213.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object ItemData
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ItemData");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ItemData", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863538.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object SubItems
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "SubItems");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "SubItems", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862053.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object DuplicateResults
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DuplicateResults");
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864553.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual NetOffice.OfficeApi.PickerFields Fields
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PickerFields>(this, "Fields", typeof(NetOffice.OfficeApi.PickerFields));
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Fields", value);
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}
