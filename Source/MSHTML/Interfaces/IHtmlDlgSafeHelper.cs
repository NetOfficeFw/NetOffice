using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IHtmlDlgSafeHelper 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface), BaseType]
	[TypeId("3050F81A-98B5-11CF-BB82-00AA00BDCE0B")]
    [CoClassSource(typeof(NetOffice.MSHTMLApi.HtmlDlgSafeHelper))]
    public interface IHtmlDlgSafeHelper : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		object fonts { get; }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		object BlockFormats { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="initColor">optional object initColor</param>
		[SupportByVersion("MSHTML", 4)]
		object choosecolordlg(object initColor);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		object choosecolordlg();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fontName">string fontName</param>
		[SupportByVersion("MSHTML", 4)]
		object getCharset(string fontName);

		#endregion
	}
}
