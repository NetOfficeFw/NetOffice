using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface IMsoDataTable 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class IMsoDataTable : COMObject, NetOffice.OfficeApi.IMsoDataTable
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
                    _type = typeof(IMsoDataTable);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IMsoDataTable() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public bool ShowLegendKey
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "ShowLegendKey");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ShowLegendKey", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public bool HasBorderHorizontal
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "HasBorderHorizontal");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HasBorderHorizontal", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public bool HasBorderVertical
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "HasBorderVertical");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HasBorderVertical", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public bool HasBorderOutline
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "HasBorderOutline");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HasBorderOutline", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.IMsoBorder Border
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoBorder>(this, "Border", typeof(NetOffice.OfficeApi.IMsoBorder));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.ChartFont Font
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ChartFont>(this, "Font", typeof(NetOffice.OfficeApi.ChartFont));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        public object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public object AutoScaleFont
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "AutoScaleFont");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "AutoScaleFont", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.IMsoChartFormat Format
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartFormat>(this, "Format", typeof(NetOffice.OfficeApi.IMsoChartFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16), ProxyResult]
        public object Application
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Application");
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public Int32 Creator
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Creator");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public void Select()
        {
            Factory.ExecuteMethod(this, "Select");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public void Delete()
        {
            Factory.ExecuteMethod(this, "Delete");
        }

        #endregion

        #pragma warning restore
    }
}
