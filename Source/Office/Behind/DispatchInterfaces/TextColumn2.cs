using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface TextColumn2 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862078.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class TextColumn2 : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.TextColumn2
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
                    _type = typeof(TextColumn2);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public TextColumn2() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860260.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public Int32 Number
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Number");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Number", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863277.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public Single Spacing
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "Spacing");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Spacing", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864646.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoTextDirection TextDirection
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTextDirection>(this, "TextDirection");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "TextDirection", value);
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}
