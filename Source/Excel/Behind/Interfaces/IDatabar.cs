using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;
 
namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface IDatabar 
    /// SupportByVersion Excel, 12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IDatabar : COMObject, NetOffice.ExcelApi.IDatabar
    {
        #pragma warning disable

        #region Type Information

        /// <summary>        /// Instance Type
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
                    _type = typeof(IDatabar);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IDatabar() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Application Application
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Enums.XlCreator Creator
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16), ProxyResult]
        public object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Int32 Priority
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Priority");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Priority", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public bool StopIfTrue
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "StopIfTrue");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Range AppliesTo
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "AppliesTo", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public NetOffice.ExcelApi.ConditionValue MinPoint
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ConditionValue>(this, "MinPoint", typeof(NetOffice.ExcelApi.ConditionValue));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public NetOffice.ExcelApi.ConditionValue MaxPoint
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ConditionValue>(this, "MaxPoint", typeof(NetOffice.ExcelApi.ConditionValue));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Int32 PercentMin
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "PercentMin");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "PercentMin", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Int32 PercentMax
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "PercentMax");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "PercentMax", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16), ProxyResult]
        public object BarColor
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "BarColor");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public bool ShowValue
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "ShowValue");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ShowValue", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Formula
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Formula");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Formula", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Int32 Type
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Type");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public bool PTCondition
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "PTCondition");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Enums.XlPivotConditionScope ScopeType
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPivotConditionScope>(this, "ScopeType");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "ScopeType", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Int32 Direction
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Direction");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Direction", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public NetOffice.ExcelApi.Enums.XlDataBarFillType BarFillType
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlDataBarFillType>(this, "BarFillType");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "BarFillType", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public NetOffice.ExcelApi.Enums.XlDataBarAxisPosition AxisPosition
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlDataBarAxisPosition>(this, "AxisPosition");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "AxisPosition", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16), ProxyResult]
        public object AxisColor
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "AxisColor");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public NetOffice.ExcelApi.DataBarBorder BarBorder
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.DataBarBorder>(this, "BarBorder", typeof(NetOffice.ExcelApi.DataBarBorder));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public NetOffice.ExcelApi.NegativeBarFormat NegativeBarFormat
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.NegativeBarFormat>(this, "NegativeBarFormat", typeof(NetOffice.ExcelApi.NegativeBarFormat));
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Int32 SetFirstPriority()
        {
            return Factory.ExecuteInt32MethodGet(this, "SetFirstPriority");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Int32 SetLastPriority()
        {
            return Factory.ExecuteInt32MethodGet(this, "SetLastPriority");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Int32 Delete()
        {
            return Factory.ExecuteInt32MethodGet(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="range">NetOffice.ExcelApi.Range range</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Int32 ModifyAppliesToRange(NetOffice.ExcelApi.Range range)
        {
            return Factory.ExecuteInt32MethodGet(this, "ModifyAppliesToRange", range);
        }

        #endregion

        #pragma warning restore
    }
}

