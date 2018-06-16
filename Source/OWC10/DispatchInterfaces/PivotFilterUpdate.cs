using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PivotFilterUpdate 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("A5E83EE4-5A92-11D3-BF58-00C04F61319A")]
	public interface PivotFilterUpdate : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="member">NetOffice.OWC10Api.PivotMember member</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum get_StateOf(NetOffice.OWC10Api.PivotMember member);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_StateOf
		/// </summary>
		/// <param name="member">NetOffice.OWC10Api.PivotMember member</param>
		[SupportByVersion("OWC10", 1), Redirect("get_StateOf")]
		NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum StateOf(NetOffice.OWC10Api.PivotMember member);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool IsDirty { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="member">NetOffice.OWC10Api.PivotMember member</param>
		[SupportByVersion("OWC10", 1)]
		void Click(NetOffice.OWC10Api.PivotMember member);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="member">NetOffice.OWC10Api.PivotMember member</param>
		/// <param name="oldMemberState">NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum oldMemberState</param>
		/// <param name="newMemberState">NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum newMemberState</param>
		[SupportByVersion("OWC10", 1)]
		void ClickFromTo(NetOffice.OWC10Api.PivotMember member, NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum oldMemberState, NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum newMemberState);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void Apply();

		#endregion
	}
}
