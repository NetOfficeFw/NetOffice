using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IElementBehaviorSiteCategory 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F4EE-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IElementBehaviorSiteCategory : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lDirection">Int32 lDirection</param>
		/// <param name="pchCategory">string pchCategory</param>
		[SupportByVersion("MSHTML", 4)]
		NetOffice.MSHTMLApi.IEnumUnknown GetRelatedBehaviors(Int32 lDirection, string pchCategory);

		#endregion
	}
}
