using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PivotFieldSet 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("F5B39AAE-1480-11D3-8549-00C04FAC67D7")]
	public interface PivotFieldSet : ICOMObject
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string Caption { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotFields Fields { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OWC10Api.PivotMembers Members { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.PivotFieldSetOrientationEnum Orientation { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.PivotFieldSetTypeEnum Type { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotField BoundField { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string UniqueName { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 Width { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		NetOffice.OWC10Api.PivotMember DefaultMember { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		NetOffice.OWC10Api.PivotMember Member { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		NetOffice.OWC10Api.PivotMember AllMember { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.PivotMembersCompareByEnum CompareOrderedMembersBy { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotView View { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OWC10Api.PivotFilterUpdate CreateFilterUpdate { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool AllowMultiFilter { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string FilterCaption { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.PivotFieldSetAllIncludeExcludeEnum AllIncludeExclude { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.PivotMembersCompareByEnum CompareMemberCaptionsBy { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool DisplayInFieldList { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool AlwaysIncludeInCube { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="nameOrPath">object nameOrPath</param>
		/// <param name="format">optional NetOffice.OWC10Api.Enums.PivotMemberFindFormatEnum format</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OWC10Api.PivotMember get_FindMember(object nameOrPath, object format);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_FindMember
		/// </summary>
		/// <param name="nameOrPath">object nameOrPath</param>
		/// <param name="format">optional NetOffice.OWC10Api.Enums.PivotMemberFindFormatEnum format</param>
		[SupportByVersion("OWC10", 1), Redirect("get_FindMember")]
		NetOffice.OWC10Api.PivotMember FindMember(object nameOrPath, object format);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="nameOrPath">object nameOrPath</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OWC10Api.PivotMember get_FindMember(object nameOrPath);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_FindMember
		/// </summary>
		/// <param name="nameOrPath">object nameOrPath</param>
		[SupportByVersion("OWC10", 1), Redirect("get_FindMember")]
		NetOffice.OWC10Api.PivotMember FindMember(object nameOrPath);

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="caption">string caption</param>
		/// <param name="dataField">string dataField</param>
		/// <param name="expression">string expression</param>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotField AddCalculatedField(string name, string caption, string dataField, string expression);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="caption">optional string Caption = </param>
		/// <param name="before">optional object Before = 0</param>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotField AddCustomGroupField(object name, object caption, object before);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotField AddCustomGroupField();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional string Name = </param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotField AddCustomGroupField(object name);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="caption">optional string Caption = </param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotField AddCustomGroupField(object name, object caption);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="field">object field</param>
		[SupportByVersion("OWC10", 1)]
		void DeleteField(object field);

		#endregion
	}
}
