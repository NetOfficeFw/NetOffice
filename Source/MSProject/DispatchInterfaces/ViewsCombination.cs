using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface ViewsCombination 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920750(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("CE4F7D83-369B-43CF-96A8-29C2DE2B8658")]
	public interface ViewsCombination : Views
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="topView">object topView</param>
		/// <param name="bottomView">object bottomView</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.ViewCombination Add(string name, object topView, object bottomView, object showInMenu);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="topView">object topView</param>
		/// <param name="bottomView">object bottomView</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.ViewCombination Add(string name, object topView, object bottomView);

		#endregion
	}
}
