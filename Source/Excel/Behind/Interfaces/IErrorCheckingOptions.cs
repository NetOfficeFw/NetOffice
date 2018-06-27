using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface IErrorCheckingOptions 
    /// SupportByVersion Excel, 10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IErrorCheckingOptions : COMObject, NetOffice.ExcelApi.IErrorCheckingOptions
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
                    _contractType = typeof(NetOffice.ExcelApi.IErrorCheckingOptions);
                return _contractType;
            }
        }
        private static Type _contractType;


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
                    _type = typeof(IErrorCheckingOptions);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IErrorCheckingOptions() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Application Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool BackgroundChecking
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "BackgroundChecking");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BackgroundChecking", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlColorIndex IndicatorColorIndex
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlColorIndex>(this, "IndicatorColorIndex");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "IndicatorColorIndex", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool EvaluateToError
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EvaluateToError");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EvaluateToError", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool TextDate
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "TextDate");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TextDate", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool NumberAsText
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "NumberAsText");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NumberAsText", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool InconsistentFormula
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "InconsistentFormula");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InconsistentFormula", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool OmittedCells
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "OmittedCells");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OmittedCells", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool UnlockedFormulaCells
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UnlockedFormulaCells");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UnlockedFormulaCells", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool EmptyCellReferences
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EmptyCellReferences");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EmptyCellReferences", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual bool ListDataValidation
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ListDataValidation");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ListDataValidation", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool InconsistentTableFormula
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "InconsistentTableFormula");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InconsistentTableFormula", value);
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}

