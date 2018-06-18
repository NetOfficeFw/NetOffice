using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IElementBehaviorFactory 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F429-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IElementBehaviorFactory : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrBehavior">string bstrBehavior</param>
		/// <param name="bstrBehaviorUrl">string bstrBehaviorUrl</param>
		/// <param name="pSite">NetOffice.MSHTMLApi.IElementBehaviorSite pSite</param>
		[SupportByVersion("MSHTML", 4)]
		NetOffice.MSHTMLApi.IElementBehavior FindBehavior(string bstrBehavior, string bstrBehaviorUrl, NetOffice.MSHTMLApi.IElementBehaviorSite pSite);

		#endregion
	}
}
