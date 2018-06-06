using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// IAutoCorrect
    /// </summary>
    [SyntaxBypass]
    public class IAutoCorrect_ : COMObject, NetOffice.ExcelApi.IAutoCorrect_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public IAutoCorrect_() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_ReplacementList(object index)
        {
            return Factory.ExecuteVariantPropertyGet(this, "ReplacementList", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_ReplacementList(object index, object value)
        {
            Factory.ExecutePropertySet(this, "ReplacementList", index, value);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_ReplacementList
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_ReplacementList")]
        public virtual object ReplacementList(object index)
        {
            return get_ReplacementList(index);
        }

        #endregion

        #region Methods

        #endregion
    }

    /// <summary>
    /// Interface IAutoCorrect 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IAutoCorrect : NetOffice.ExcelApi.Behind.IAutoCorrect_, NetOffice.ExcelApi.IAutoCorrect
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
                    _type = typeof(IAutoCorrect);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public IAutoCorrect() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Application Application
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
        public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
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
        public virtual object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CapitalizeNamesOfDays
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "CapitalizeNamesOfDays");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "CapitalizeNamesOfDays", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ReplacementList
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "ReplacementList");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "ReplacementList", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ReplaceText
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "ReplaceText");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ReplaceText", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool TwoInitialCapitals
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "TwoInitialCapitals");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "TwoInitialCapitals", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CorrectSentenceCap
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "CorrectSentenceCap");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "CorrectSentenceCap", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CorrectCapsLock
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "CorrectCapsLock");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "CorrectCapsLock", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool DisplayAutoCorrectOptions
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "DisplayAutoCorrectOptions");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "DisplayAutoCorrectOptions", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual bool AutoExpandListRange
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "AutoExpandListRange");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "AutoExpandListRange", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool AutoFillFormulasInLists
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "AutoFillFormulasInLists");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "AutoFillFormulasInLists", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">string what</param>
        /// <param name="replacement">string replacement</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AddReplacement(string what, string replacement)
        {
            return Factory.ExecuteVariantMethodGet(this, "AddReplacement", what, replacement);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">string what</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object DeleteReplacement(string what)
        {
            return Factory.ExecuteVariantMethodGet(this, "DeleteReplacement", what);
        }

        #endregion

        #pragma warning restore
    }
}

