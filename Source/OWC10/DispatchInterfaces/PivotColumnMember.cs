using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PivotColumnMember 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("1D40A585-EBA2-11D2-8F35-00600893B533")]
	public interface PivotColumnMember : PivotAxisMember
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotColumnMembers ChildColumnMembers { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotColumnMember ParentColumnMember { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="path">string path</param>
		/// <param name="format">NetOffice.OWC10Api.Enums.PivotMemberFindFormatEnum format</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OWC10Api.PivotColumnMember get_FindColumnMember(string path, NetOffice.OWC10Api.Enums.PivotMemberFindFormatEnum format);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_FindColumnMember
		/// </summary>
		/// <param name="path">string path</param>
		/// <param name="format">NetOffice.OWC10Api.Enums.PivotMemberFindFormatEnum format</param>
		[SupportByVersion("OWC10", 1), Redirect("get_FindColumnMember")]
		NetOffice.OWC10Api.PivotColumnMember FindColumnMember(string path, NetOffice.OWC10Api.Enums.PivotMemberFindFormatEnum format);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotColumnMember TotalColumnMember { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 DetailLeft { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 DetailLeftOffset { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool DetailsExpanded { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="detailLeft">Int32 detailLeft</param>
		/// <param name="detailLeftOffset">Int32 detailLeftOffset</param>
		/// <param name="update">optional bool Update = true</param>
		[SupportByVersion("OWC10", 1)]
		void MoveDetailLeft(Int32 detailLeft, Int32 detailLeftOffset, object update);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="detailLeft">Int32 detailLeft</param>
		/// <param name="detailLeftOffset">Int32 detailLeftOffset</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void MoveDetailLeft(Int32 detailLeft, Int32 detailLeftOffset);

		#endregion
	}
}
