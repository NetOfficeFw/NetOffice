using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface IFormatCondition 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IFormatCondition : COMObject, NetOffice.ExcelApi.IFormatCondition
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
                    _type = typeof(IFormatCondition);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IFormatCondition() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Application Application
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Enums.XlCreator Creator
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Type
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Type");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Operator
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Operator");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Formula1
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Formula1");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Formula2
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Formula2");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Interior Interior
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Interior>(this, "Interior", typeof(NetOffice.ExcelApi.Interior));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Borders Borders
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Borders>(this, "Borders", typeof(NetOffice.ExcelApi.Borders));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Font Font
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Font>(this, "Font", typeof(NetOffice.ExcelApi.Font));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Text
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Text");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Text", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Enums.XlContainsOperator TextOperator
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlContainsOperator>(this, "TextOperator");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "TextOperator", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Enums.XlTimePeriods DateOperator
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTimePeriods>(this, "DateOperator");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "DateOperator", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public object NumberFormat
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "NumberFormat");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "NumberFormat", value);
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
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public bool StopIfTrue
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "StopIfTrue");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "StopIfTrue", value);
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

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="formula1">optional object formula1</param>
        /// <param name="formula2">optional object formula2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2)
        {
            return Factory.ExecuteInt32MethodGet(this, "Modify", type, _operator, formula1, formula2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="formula1">optional object formula1</param>
        /// <param name="formula2">optional object formula2</param>
        /// <param name="_string">optional object string</param>
        /// <param name="operator2">optional object operator2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Int32 Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2, object _string, object operator2)
        {
            return Factory.ExecuteInt32MethodGet(this, "Modify", new object[] { type, _operator, formula1, formula2, _string, operator2 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type)
        {
            return Factory.ExecuteInt32MethodGet(this, "Modify", type);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator)
        {
            return Factory.ExecuteInt32MethodGet(this, "Modify", type, _operator);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="formula1">optional object formula1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1)
        {
            return Factory.ExecuteInt32MethodGet(this, "Modify", type, _operator, formula1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="formula1">optional object formula1</param>
        /// <param name="formula2">optional object formula2</param>
        /// <param name="_string">optional object string</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Int32 Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2, object _string)
        {
            return Factory.ExecuteInt32MethodGet(this, "Modify", new object[] { type, _operator, formula1, formula2, _string });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Delete()
        {
            return Factory.ExecuteInt32MethodGet(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="formula1">optional object formula1</param>
        /// <param name="formula2">optional object formula2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Int32 _Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2)
        {
            return Factory.ExecuteInt32MethodGet(this, "_Modify", type, _operator, formula1, formula2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Int32 _Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type)
        {
            return Factory.ExecuteInt32MethodGet(this, "_Modify", type);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Int32 _Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator)
        {
            return Factory.ExecuteInt32MethodGet(this, "_Modify", type, _operator);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="formula1">optional object formula1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Int32 _Modify(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1)
        {
            return Factory.ExecuteInt32MethodGet(this, "_Modify", type, _operator, formula1);
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

        #endregion

        #pragma warning restore
    }
}

