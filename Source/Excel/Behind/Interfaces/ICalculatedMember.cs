using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface ICalculatedMember 
    /// SupportByVersion Excel, 10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class ICalculatedMember : COMObject, NetOffice.ExcelApi.ICalculatedMember
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
                    _type = typeof(ICalculatedMember);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ICalculatedMember() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Application Application
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Enums.XlCreator Creator
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16), ProxyResult]
        public object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public string Name
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Name");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public string Formula
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Formula");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public string SourceName
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "SourceName");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public Int32 SolveOrder
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "SolveOrder");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public bool IsValid
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "IsValid");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public string _Default
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "_Default");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Enums.XlCalculatedMemberType Type
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCalculatedMemberType>(this, "Type");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public bool Dynamic
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "Dynamic");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public string DisplayFolder
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "DisplayFolder");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public bool HierarchizeDistinct
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "HierarchizeDistinct");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HierarchizeDistinct", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public bool FlattenHierarchies
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "FlattenHierarchies");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "FlattenHierarchies", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        public string MeasureGroup
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "MeasureGroup");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        public string ParentHierarchy
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "ParentHierarchy");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        public string ParentMember
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "ParentMember");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        public NetOffice.ExcelApi.Enums.XlCalcMemNumberFormatType NumberFormat
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCalcMemNumberFormatType>(this, "NumberFormat");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public Int32 Delete()
        {
            return Factory.ExecuteInt32MethodGet(this, "Delete");
        }

        #endregion

        #pragma warning restore
    }
}

