using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IElementBehaviorSite 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("3050F427-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IElementBehaviorSite : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLElement GetElement();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lEvent">Int32 lEvent</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 RegisterNotification(Int32 lEvent);

		#endregion
	}
}
