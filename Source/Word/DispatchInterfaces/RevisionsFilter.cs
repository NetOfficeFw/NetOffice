using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface RevisionsFilter 
	/// SupportByVersion Word, 15, 16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228191.aspx </remarks>
	[SupportByVersion("Word", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("D523C26B-7278-4FA9-AA0B-0827DC8B41CE")]
	public interface RevisionsFilter : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231521.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.WordApi.Enums.WdRevisionsView View { get; set; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230964.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.WordApi.Enums.WdRevisionsMarkup Markup { get; set; }

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231708.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		NetOffice.WordApi.Reviewers Reviewers { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227337.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		void ToggleShowAllReviewers();

		#endregion
	}
}
