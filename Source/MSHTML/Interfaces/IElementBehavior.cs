using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IElementBehavior 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F425-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IElementBehavior : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pBehaviorSite">NetOffice.MSHTMLApi.IElementBehaviorSite pBehaviorSite</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 Init(NetOffice.MSHTMLApi.IElementBehaviorSite pBehaviorSite);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lEvent">Int32 lEvent</param>
		/// <param name="pVar">object pVar</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 Notify(Int32 lEvent, object pVar);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 Detach();

		#endregion
	}
}
