using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
    /// <summary>
    /// RecordsetDef
    /// </summary>
    [SyntaxBypass]
    public interface RecordsetDef_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="fetchType">optional NetOffice.OWC10Api.Enums.DscFetchTypeEnum fetchType</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_ShapeText(object fetchType);

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_ShapeText
        /// </summary>
        /// <param name="fetchType">optional NetOffice.OWC10Api.Enums.DscFetchTypeEnum fetchType</param>
        [SupportByVersion("OWC10", 1), Redirect("get_ShapeText")]
        string ShapeText(object fetchType);

        #endregion
    }

    /// <summary>
    /// DispatchInterface RecordsetDef 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("F5B39A9D-1480-11D3-8549-00C04FAC67D7")]
    public interface RecordsetDef : RecordsetDef_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        string Name { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        new string ShapeText { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        string CommandText { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool DataEntry { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        string ServerFilter { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.PageRowsource PrimaryPageRowsource { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.SublistRelationships SublistRelationships { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.PageFields PageFields { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.RecordsetDef ParentRecordsetDef { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.GroupingDefs GroupingDefs { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.ParameterValues ParameterValues { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.PageRowsources PageRowsources { get; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        string UniqueTable { get; set; }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        string ResyncCommand { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        NetOffice.OWC10Api.RecordsetDef Demote();

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        void Delete();

        #endregion
    }
}
