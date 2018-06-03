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
        public string Id
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Id");
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861059.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public string DisplayName
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "DisplayName");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "DisplayName", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865231.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public string Type
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Type");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Type", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863831.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public string SIPId
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "SIPId");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "SIPId", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865213.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public object ItemData
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "ItemData");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "ItemData", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863538.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public object SubItems
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "SubItems");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "SubItems", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862053.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public object DuplicateResults
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "DuplicateResults");
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864553.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public NetOffice.OfficeApi.PickerFields Fields
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PickerFields>(this, "Fields", typeof(NetOffice.OfficeApi.PickerFields));
            }
            set
            {
                Factory.ExecuteReferencePropertySet(this, "Fields", value);
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}
