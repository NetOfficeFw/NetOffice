using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface IColorScale 
    /// SupportByVersion Excel, 12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IColorScale : COMObject, NetOffice.ExcelApi.IColorScale
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
                    _type = typeof(IColorScale);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IColorScale() : base()
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
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public NetOffice.ExcelApi.ColorScaleCriteria ColorScaleCriteria
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ColorScaleCriteria>(this, "ColorScaleCriteria", typeof(NetOffice.ExcelApi.ColorScaleCriteria));
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

