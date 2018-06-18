using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IElementBehaviorSiteOM 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface), BaseType]
	[TypeId("3050F489-98B5-11CF-BB82-00AA00BDCE0B")]
	public interface IElementBehaviorSiteOM : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pchEvent">string pchEvent</param>
		/// <param name="lFlags">Int32 lFlags</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 RegisterEvent(string pchEvent, Int32 lFlags);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pchEvent">string pchEvent</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetEventCookie(string pchEvent);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lCookie">Int32 lCookie</param>
		/// <param name="pEventObject">NetOffice.MSHTMLApi.IHTMLEventObj pEventObject</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 FireEvent(Int32 lCookie, NetOffice.MSHTMLApi.IHTMLEventObj pEventObject);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		NetOffice.MSHTMLApi.IHTMLEventObj CreateEventObject();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pchName">string pchName</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 RegisterName(string pchName);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pchUrn">string pchUrn</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 RegisterUrn(string pchUrn);

		#endregion
	}
}
