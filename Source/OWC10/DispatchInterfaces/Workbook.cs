using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
    /// <summary>
    /// Workbook
    /// </summary>
    [SyntaxBypass]
    public interface Workbook_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_Colors(object index);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_Colors(object index, object value);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_Colors
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("OWC10", 1), Redirect("get_Colors")]
        object Colors(object index);

        #endregion
    }

    /// <summary>
    /// DispatchInterface Workbook 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("F5B39BA6-1480-11D3-8549-00C04FAC67D7")]
    public interface Workbook : Workbook_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.Worksheet ActiveSheet { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        NetOffice.OWC10Api.ISpreadsheet Application { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        Int32 CalculationVersion { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        new object Colors { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        string Name { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.Names Names { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [BaseResult]
        NetOffice.OWC10Api.ISpreadsheet Parent { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        bool ProtectStructure { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.Sheets Sheets { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.Windows Windows { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.Worksheets Worksheets { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="structure">optional object structure</param>
        /// <param name="windows">optional object windows</param>
        [SupportByVersion("OWC10", 1)]
        void Protect(object password, object structure, object windows);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void Protect();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="password">optional object password</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void Protect(object password);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="structure">optional object structure</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void Protect(object password, object structure);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        void ResetColors();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="password">optional object password</param>
        [SupportByVersion("OWC10", 1)]
        void Unprotect(object password);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        void Unprotect();

        #endregion
    }
}
