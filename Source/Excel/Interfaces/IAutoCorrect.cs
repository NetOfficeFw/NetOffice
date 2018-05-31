using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// IAutoCorrect
    /// </summary>
    [SyntaxBypass]
    public interface IAutoCorrect_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_ReplacementList(object index);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_ReplacementList(object index, object value);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_ReplacementList
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_ReplacementList")]
        object ReplacementList(object index);

        #endregion
    }

    /// <summary>
    /// Interface IAutoCorrect 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public interface IAutoCorrect : IAutoCorrect_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool CapitalizeNamesOfDays { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object ReplacementList { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool ReplaceText { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool TwoInitialCapitals { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool CorrectSentenceCap { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        bool CorrectCapsLock { get; set; }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        bool DisplayAutoCorrectOptions { get; set; }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        bool AutoExpandListRange { get; set; }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        bool AutoFillFormulasInLists { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">string what</param>
        /// <param name="replacement">string replacement</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object AddReplacement(string what, string replacement);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">string what</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        object DeleteReplacement(string what);

        #endregion
    }
}
