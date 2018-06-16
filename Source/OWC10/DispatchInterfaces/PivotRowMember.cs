using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PivotRowMember 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("1D40A584-EBA2-11D2-8F35-00600893B533")]
	public interface PivotRowMember : PivotAxisMember
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotRowMembers ChildRowMembers { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotRowMember ParentRowMember { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="path">string path</param>
		/// <param name="format">NetOffice.OWC10Api.Enums.PivotMemberFindFormatEnum format</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OWC10Api.PivotRowMember get_FindRowMember(string path, NetOffice.OWC10Api.Enums.PivotMemberFindFormatEnum format);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_FindRowMember
		/// </summary>
		/// <param name="path">string path</param>
		/// <param name="format">NetOffice.OWC10Api.Enums.PivotMemberFindFormatEnum format</param>
		[SupportByVersion("OWC10", 1), Redirect("get_FindRowMember")]
		NetOffice.OWC10Api.PivotRowMember FindRowMember(string path, NetOffice.OWC10Api.Enums.PivotMemberFindFormatEnum format);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.PivotRowMember TotalRowMember { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 TotalRowHeight { get; set; }

		#endregion

	}
}
